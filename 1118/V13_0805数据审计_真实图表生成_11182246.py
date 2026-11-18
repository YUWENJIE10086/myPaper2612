#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
V13 / 0805多规格按日查询旧版xls实证审计器（只读，不改原始文件）
运行: pip install xlrd matplotlib
      python V13_0805数据审计_真实图表生成_11182246.py --repo-root .
输出: 1118/V13_本地审计结果_不自动提交/  (含 JSON、CSV、SVG、PNG)
注意: 此脚本只做原始数据核验、真实EDA与季节性基线；
      不会假装已经训练TFT、GAT或MTGNN。
"""
import argparse, csv, json, math, re, statistics, calendar
from collections import defaultdict, Counter
from datetime import datetime, date, timedelta
from pathlib import Path

try:
    import xlrd
except ImportError:
    raise SystemExit("缺少xls读取依赖；运行 pip install xlrd matplotlib")

QTY_WORDS = ("销量", "销售数量", "销售量", "数量", "购进量", "订购量", "销售条数", "订购条数", "日销量")
SKU_WORDS = ("规格编码", "品规编码", "商品编码", "商品代码", "卷烟编码", "规格名称", "品规名称", "商品名称", "卷烟名称", "品牌规格", "规格", "品规", "商品")
DATE_WORDS = ("日期", "交易日期", "销售日期", "统计日期", "业务日期", "时间")
BLOCK_WORDS = ("合计", "汇总", "总计", "小计", "平均", "整体")

def norm(v):
    return re.sub(r"\s+", "", str(v or "")).lower()

def clean_text(v):
    v = str(v).strip() if v is not None else ""
    return re.sub(r"\s+", " ", v).strip()

def as_number(v):
    if isinstance(v, bool) or v is None or v == "":
        return None
    try:
        if isinstance(v, str):
            v = v.strip().replace(",", "").replace("，", "")
            if not v or v in ("--", "-", "/", "—"):
                return None
        n = float(v)
        return n if math.isfinite(n) else None
    except (ValueError, TypeError):
        return None

def header_match(v, keywords):
    s = norm(v)
    return any(x in s for x in keywords) if s else False

def infer_date(v, cell_type, datemode, yearmonth, day_only_ok=False):
    year, month = int(yearmonth[:4]), int(yearmonth[4:6])
    if cell_type == xlrd.XL_CELL_DATE:
        try: return xlrd.xldate_as_datetime(v, datemode).date()
        except Exception: return None
    s = clean_text(v).replace("年","-").replace("月","-").replace("日","").replace("号","")
    if re.fullmatch(r"20\d{6}", s):
        try: return datetime.strptime(s, "%Y%m%d").date()
        except ValueError: return None
    if re.fullmatch(r"20\d{2}[-/]\d{1,2}[-/]\d{1,2}", s):
        try: return date.fromisoformat(s.replace("/","-"))
        except ValueError: return None
    num = as_number(v)
    if num is not None and num >= 40000:
        try: return xlrd.xldate_as_datetime(num, datemode).date()
        except Exception: pass
    if day_only_ok and num is not None and int(num) == num and 1 <= num <= 31:
        try: return date(year, month, int(num))
        except ValueError: return None
    if day_only_ok and re.fullmatch(r"\d{1,2}", s):
        try: return date(year,month,int(s))
        except ValueError: return None
    return None

def get(sheet,r,c):
    if r<0 or c<0 or r>=sheet.nrows or c>=sheet.ncols: return ""
    return sheet.cell_value(r,c)

def collect_workbook(path, diagnostics):
    ym=path.stem
    book=xlrd.open_workbook(str(path), on_demand=True)
    results=[]
    for sh in book.sheets():
        diag={"file":path.name,"sheet":sh.name,"nrows":sh.nrows,"ncols":sh.ncols,"mode":"not_detected"}
        diagnostics.append(diag)
        # The xls sheets may be wide: one row per SKU, day-of-month as columns.
        heads=[]
        for ri in range(min(sh.nrows,14)):
            daycols=[]
            for ci in range(sh.ncols):
                d=infer_date(get(sh,ri,ci),sh.cell_type(ri,ci),book.datemode,ym,True)
                if d and d.strftime("%Y%m")==ym:
                    daycols.append((ci,d))
            if len({d for _,d in daycols})>=7:
                heads.append((ri,daycols))
        if heads:
            hr,ds=max(heads,key=lambda z:len({d for _,d in z[1]}))
            min_day_col=min(i for i,_ in ds)
            sku_cols=[]
            for sr in range(max(0,hr-2),hr+1):
                for ci in range(min_day_col):
                    v=get(sh,sr,ci)
                    if header_match(v,SKU_WORDS):
                        sku_cols.append((ci,len(norm(v))))
            sku_col=sorted(sku_cols,key=lambda x:(x[1],-x[0]),reverse=True)[0][0] if sku_cols else None
            if sku_col is None:
                diag["mode"]="wide_header_detected_sku_not_detected"
                diag["day_columns"]=len(ds)
                continue
            diag.update({"mode":"wide","header_row":hr+1,"sku_col":sku_col+1,"day_columns":len(ds)})
            for ri in range(hr+1,sh.nrows):
                sku=clean_text(get(sh,ri,sku_col))
                if not sku or header_match(sku,BLOCK_WORDS):
                    continue
                for ci,d in ds:
                    n=as_number(get(sh,ri,ci))
                    if n is not None:
                        results.append((d.isoformat(),sku,n,path.name,sh.name))
            continue
        # Long form: one row per date×sku.
        detected=False
        for hr in range(min(sh.nrows,14)):
            vals=[get(sh,hr,c) for c in range(sh.ncols)]
            skus=[i for i,v in enumerate(vals) if header_match(v,SKU_WORDS)]
            dts=[i for i,v in enumerate(vals) if header_match(v,DATE_WORDS)]
            qtys=[i for i,v in enumerate(vals) if header_match(v,QTY_WORDS)]
            qtys=[i for i in qtys if i not in skus+dts]
            if skus and dts and qtys:
                si,di,qi=skus[0],dts[0],qtys[0]
                diag.update({"mode":"long","header_row":hr+1,"sku_col":si+1,"date_col":di+1,"qty_col":qi+1})
                for ri in range(hr+1,sh.nrows):
                    d=infer_date(get(sh,ri,di),sh.cell_type(ri,di),book.datemode,ym)
                    sku=clean_text(get(sh,ri,si))
                    n=as_number(get(sh,ri,qi))
                    if d and sku and n is not None and not header_match(sku,BLOCK_WORDS):
                        results.append((d.isoformat(),sku,n,path.name,sh.name))
                detected=True
                break
        if not detected:
            diag["mode"]="unidentified"
            diag["header_snapshot"]=[
                [clean_text(get(sh,r,c))[:38] for c in range(min(12,sh.ncols))]
                for r in range(min(7,sh.nrows))
            ]
    return results

def csv_dump(path,heads,rows):
    with path.open("w",encoding="utf-8-sig",newline="") as f:
        w=csv.writer(f)
        w.writerow(heads)
        w.writerows(rows)

def fnum(v):
    return round(float(v),6) if v is not None else None

def baseline(predictions, skus, dates, data, valstart, teststart, testend):
    # Only two transparent, legitimate P1 baselines; no TFT/MTGNN results invented.
    result=[]
    for sku in skus:
        series={d:data[(d,sku)] for d in dates if (d,sku) in data}
        for day in dates:
            if day < teststart or day > testend or day not in series:
                continue
            last7=[(day-timedelta(days=k)).isoformat() for k in range(1,8)]
            if not all(d in series for d in last7):
                continue
            pred={"Lag7":series[last7[-1]],"Mean7":sum(series[d] for d in last7)/7}
            y=series[day]
            for label,p in pred.items():
                result.append((day.isoformat(),sku,label,y,p,abs(y-p)))
    return result

def metrics_for(rows):
    if not rows:return []
    names=sorted(set(x[2] for x in rows))
    metrics=[]
    for method in names:
        rr=[x for x in rows if x[2]==method]
        n=len(rr)
        mae=sum(x[5] for x in rr)/n
        real=sum(abs(x[3]) for x in rr)
        wmape=sum(x[5] for x in rr)/real if real>0 else None
        metrics.append((method,n,fnum(mae),fnum(wmape)))
    return metrics

def plot_graphs(out, inventory, data, dates, skus, scored):
    try:
        import matplotlib
        matplotlib.use("Agg")
        import matplotlib.pyplot as plt
        from matplotlib.ticker import MaxNLocator
    except ImportError:
        print("未安装matplotlib，跳过图表；pip install matplotlib")
        return []
    plt.rcParams.update({"font.size":10,"axes.unicode_minus":False})
    names=[x[0][2:] for x in inventory]
    sizes=[x[1]/1024 for x in inventory]
    # Figure D0: source-file sizes are actual metadata, not rows/retail sales.
    fig,ax=plt.subplots(figsize=(12,4.2))
    ax.plot(names,sizes,marker="o",lw=1.9)
    ax.set_title("D0. File-size audit of 19 monthly .xls sources (not sales)")
    ax.set_xlabel("Source month (YYMM)");ax.set_ylabel("Git blob bytes / KiB")
    ax.tick_params(axis="x",rotation=50,labelsize=8)
    ax.grid(alpha=.18)
    fig.tight_layout();fig.savefig(out/"D0_monthly_blob_size.svg");fig.savefig(out/"D0_monthly_blob_size.png",dpi=300);plt.close(fig)
    figures=["D0_monthly_blob_size"]
    if not data:return figures
    total=defaultdict(float);cov=Counter()
    for (d,s),v in data.items():
        total[d]+=v;cov[d]+=1
    sorted_days=sorted(total)
    fig,ax=plt.subplots(figsize=(12,4))
    ax.plot(sorted_days,[cov[x] for x in sorted_days],lw=1.25)
    ax.set_title("D1. SKU coverage per observed date (missing is not zero)")
    ax.set_xlabel("Observed calendar date");ax.set_ylabel("Observed SKU count")
    ax.xaxis.set_major_locator(MaxNLocator(9));ax.tick_params(axis="x",rotation=25)
    fig.tight_layout();fig.savefig(out/"D1_coverage.svg");fig.savefig(out/"D1_coverage.png",dpi=300);plt.close(fig)
    figures.append("D1_coverage")
    fig,ax=plt.subplots(figsize=(12,4))
    ax.plot(sorted_days,[total[x] for x in sorted_days],lw=1.1)
    ax.set_title("D2. Raw aggregate of observed quantity (verify measurement unit)")
    ax.set_xlabel("Observed date");ax.set_ylabel("Sum of measured raw values")
    ax.xaxis.set_major_locator(MaxNLocator(9));ax.tick_params(axis="x",rotation=25)
    fig.tight_layout();fig.savefig(out/"D2_aggregate_observed.svg");fig.savefig(out/"D2_aggregate_observed.png",dpi=300);plt.close(fig)
    figures.append("D2_aggregate_observed")
    if scored:
        mt=metrics_for(scored)
        fig,ax=plt.subplots(figsize=(7.5,4))
        x=[r[0] for r in mt]
        y=[100*r[3] if r[3] is not None else 0 for r in mt]
        ax.bar(x,y)
        ax.set_title("P1. Holdout wMAPE, seasonal baselines only")
        ax.set_ylabel("wMAPE (%)")
        fig.tight_layout();fig.savefig(out/"P1_baseline_wmape.svg");fig.savefig(out/"P1_baseline_wmape.png",dpi=300);plt.close(fig)
        figures.append("P1_baseline_wmape")
    return figures

def main():
    ap=argparse.ArgumentParser()
    ap.add_argument("--repo-root",default=".")
    args=ap.parse_args()
    root=Path(args.repo_root).resolve()
    folder=root/"0805数据"/"多规格按日查询"
    out=root/"1118"/"V13_本地审计结果_不自动提交"
    out.mkdir(parents=True,exist_ok=True)
    files=sorted(folder.glob("20????.xls"))
    if not files:
        raise SystemExit("未找到19份源文件；请使用完整仓库根路径传入 --repo-root")
    inv=[(p.stem,p.stat().st_size) for p in files]
    diagnostics=[];raw=[]
    for path in files:
        try:raw.extend(collect_workbook(path,diagnostics))
        except Exception as e:diagnostics.append({"file":path.name,"error":type(e).__name__+": "+str(e)})
    csv_dump(out/"00_源文件体量.csv",["month","source_bytes"],inv)
    # The stored file may be a cross-tab with subtotals / different quantity
    # semantics. On detected duplicate (date,SKU), stop rather than sum silently.
    values=defaultdict(list)
    for d,sku,v,file,sheet in raw:values[(d,sku)].append((v,file,sheet))
    duplicates={k:v for k,v in values.items() if len(v)!=1}
    conflicted=len(duplicates)
    data={k:v[0][0] for k,v in values.items() if len(v)==1}
    skipped_duplicates=sum(len(v) for v in duplicates.values())
    valid_date=sorted(set(date.fromisoformat(k[0]) for k in data))
    skus=sorted(set(k[1] for k in data))
    csv_dump(out/"01_解析诊断.csv",
        ["file","sheet","mode","nrows","ncols","header_row","sku_col","day_columns","date_col","qty_col","error"],
        [[d.get(k,"") for k in ("file","sheet","mode","nrows","ncols","header_row","sku_col","day_columns","date_col","qty_col","error")] for d in diagnostics])
    # Raw detailed records remain only in LOCAL output; never auto-commit.
    csv_dump(out/"02_SKU日度明细_仅本地.csv",["date","sku","measured_quantity"],[(d,s,v) for (d,s),v in sorted(data.items())])
    coverage=Counter(d for d,s in data)
    bysku=Counter(s for d,s in data)
    teststart=date(2026,3,1); testend=date(2026,7,31)
    scored=baseline([],skus,valid_date,data,date(2026,1,1),teststart,testend)
    mt=metrics_for(scored)
    csv_dump(out/"03_季节基线成绩_仅独立留出.csv",["model","paired_test_rows","MAE","wMAPE"],mt)
    figures=plot_graphs(out,inv,data,valid_date,skus,scored)
    result={
        "source_glob":str(folder/"20????.xls"),
        "source_files":len(files),
        "month_range":[min(x[0] for x in inv),max(x[0] for x in inv)],
        "source_total_bytes":sum(x[1] for x in inv),
        "sheets_total":len(diagnostics),
        "detected_sheets":Counter(d.get("mode","error") for d in diagnostics),
        "parsed_original_cells":len(raw),
        "unique_date_sku_rows_after_dedup_exclusion":len(data),
        "duplicate_keys_excluded":conflicted,
        "duplicate_raw_cells_excluded":skipped_duplicates,
        "sku_count":len(skus),
        "observed_dates":len(valid_date),
        "min_date":str(min(valid_date)) if valid_date else None,
        "max_date":str(max(valid_date)) if valid_date else None,
        "median_observed_sku_per_date":statistics.median(coverage.values()) if coverage else None,
        "median_observed_dates_per_sku":statistics.median(bysku.values()) if bysku else None,
        "seasonal_baseline":mt,
        "figures":figures,
        "limitations":[
            "Only exact date-SKU keys with one source value are used; repeated keys excluded, NOT summed.",
            "Zero sales vs no observation vs out-of-stock are NOT inferable from this raw format.",
            "Parsing is heuristic and MUST be manually verified against at least 2 original xls sheets.",
            "Original measurement unit and actual demand vs realized sales are UNKNOWN.",
            "No TFT/MTGNN/GAT is trained; no model-specific numeric results can be claimed."
        ]
    }
    with (out/"04_汇总审计.json").open("w",encoding="utf-8") as f:
        json.dump(result,f,ensure_ascii=False,indent=2,default=dict)
    print(json.dumps(result,ensure_ascii=False,indent=2,default=dict))
    print("\n注意：未经原表核对统计口径，不得将本脚本生成的数值写入投稿论文。")

if __name__=="__main__":
    main()

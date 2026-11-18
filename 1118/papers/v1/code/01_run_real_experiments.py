#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
V1 REAL EMPIRICAL EXPERIMENTS, NO FABRICATED NUMBERS
Source 1: raw, original 0805数据/*.xls on the authenticated GitHub runner
Source 2: previous auditable derived table Papers2/v18/data/raw_panel_v1.csv
Task: SKU x month realized volume forecasting. (NOT SKU x day.)
Model baselines: lag-1, past-3 mean, lag-12, Ridge, HistGB, and graph-informed HistGB.
The graph is fitted only on data before each forecast origin; predictions do NOT
establish causal consumer switching, and do not claim to be TFT/MTGNN.
Run in repository root: python 1118/papers/v1/code/01_run_real_experiments.py
"""
import os, re, json, hashlib, math, warnings, calendar
from pathlib import Path
from collections import Counter, defaultdict
import numpy as np
import pandas as pd
from sklearn.ensemble import HistGradientBoostingRegressor
from sklearn.linear_model import Ridge
from sklearn.preprocessing import StandardScaler
from sklearn.pipeline import make_pipeline
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

ROOT=Path(__file__).resolve().parents[4]
OUT=ROOT/"1118/papers/v1"
FIG=OUT/"figures"
RES=OUT/"results"
FIG.mkdir(parents=True,exist_ok=True)
RES.mkdir(parents=True,exist_ok=True)
np.random.seed(20261118)
plt.rcParams.update({"font.size":10,"axes.unicode_minus":False,"svg.fonttype":"none"})
COLORS={"naive":"#8794a8","ridge":"#8f6fbe","tree":"#1c78a6","graph":"#118c86"}
MONTHS=pd.period_range("2025-01","2026-07",freq="M")
MN=[p.strftime("%Y%m") for p in MONTHS]
IDX={m:i for i,m in enumerate(MN)}

def num(x):
    try:return float(x)
    except:return np.nan

def sku(x):
    if pd.isna(x):return ""
    s=str(x).strip().replace("\u00a0","")
    if s.endswith(".0") and s[:-2].isdigit():s=s[:-2]
    if re.fullmatch(r"\d{13}",s):return s
    try:
        n=float(s)
        if n.is_integer() and len(str(int(n)))==13:return str(int(n))
    except:pass
    return ""

def output_json(p,v):
    p.write_text(json.dumps(v,ensure_ascii=False,indent=2,default=lambda x:x.item() if isinstance(x,np.generic) else str(x)),encoding="utf-8")

def source_audit():
    p=ROOT/"0805数据"
    raw=[]
    for sub in ["多规格按日查询","多指标销售汇总","品牌月进货情况查询"]:
        for f in sorted((p/sub).glob("20????.xls")):
            record={"group":sub,"month":f.stem,"bytes":f.stat().st_size,"sheets":[]}
            try:
                import xlrd
                b=xlrd.open_workbook(str(f),on_demand=True)
                for s in b.sheets():
                    record["sheets"].append({"name":s.name,"rows":s.nrows,"cols":s.ncols})
                b.release_resources()
            except Exception as e:
                record["error"]=type(e).__name__+": "+str(e)
            raw.append(record)
    output_json(RES/"00_original_xls_structure.json",raw)
    # Verify a direct field mapping from the earlier builder script.
    original={}
    dups=0
    bad=0
    for f in sorted((p/"多指标销售汇总").glob("20????.xls")):
        try:
            import xlrd
            b=xlrd.open_workbook(str(f),on_demand=True)
            for sh in b.sheets():
                for i in range(sh.nrows):
                    if sh.ncols <= 10:continue
                    code=sku(sh.cell_value(i,3))
                    if not code:continue
                    value=num(sh.cell_value(i,10))
                    key=(f.stem,code)
                    if key in original:dups+=1
                    elif np.isfinite(value):original[key]=value
                    else:bad+=1
            b.release_resources()
        except Exception as e:
            raw.append({"file":str(f),"fatal_original_check_error":str(e)})
    return raw,original,dups,bad

def load_panel():
    p=ROOT/"Papers2/v18/data/raw_panel_v1.csv"
    x=pd.read_csv(p,dtype={"ym":str,"code":str},low_memory=False)
    x["code"]=x["code"].map(sku)
    x["ym"]=x["ym"].astype(str)
    x["vol"]=pd.to_numeric(x["vol"],errors="coerce")
    x["cat"]=x["cat"].fillna("未知").astype(str)
    x=x[x["code"].ne("")&x["ym"].isin(MN)].copy()
    assert not x.duplicated(["ym","code"]).any(), "Duplicate month-SKU keys, cannot silently sum"
    return x

def verify_original(x,original,dups,bad):
    obs=x[x.vol.notna()].copy()
    obs["direct"]=obs.apply(lambda z: original.get((str(z.ym),str(z.code)),np.nan),axis=1)
    matched=obs.direct.notna()
    absdiff=(obs.loc[matched,"vol"]-obs.loc[matched,"direct"]).abs()
    return {
        "direct_original_sales_records":len(original),
        "derived_panel_rows":len(x),
        "derived_vol_nonmissing":len(obs),
        "source_matched_rows":int(matched.sum()),
        "source_matched_fraction":round(float(matched.mean()),6) if len(matched) else None,
        "source_exact_fraction_of_matched":round(float((absdiff<1e-6).mean()),6) if len(absdiff) else None,
        "source_max_abs_delta":round(float(absdiff.max()),6) if len(absdiff) else None,
        "original_duplicate_month_sku":dups,
        "original_unreadable_volume_cells":bad,
    }

def prepare_matrix(x):
    codes=sorted(x.code.unique())
    keys={c:i for i,c in enumerate(codes)}
    mat=np.full((len(codes),len(MN)),np.nan)
    cat={}
    metadata={}
    for z in x.itertuples():
        i=keys[z.code]
        mat[i,IDX[z.ym]]=z.vol if pd.notna(z.vol) and float(z.vol)>=0 else np.nan
        if z.code not in cat:cat[z.code]=str(z.cat)
        metadata[z.code]=str(getattr(z,"name",""))
    return codes,mat,cat,metadata

def numeric_features(v,i,t,h,peer):
    a,b,c=v[i,t-1],v[i,t-2],v[i,t-3]
    if not all(np.isfinite(q) and q>=0 for q in (a,b,c)):return None
    la,lb,lc=np.log1p([a,b,c])
    target_p=MONTHS[t+h-1]
    period=int(target_p.month)
    # NO current/future demand, price, inventory or fulfillment allowed.
    return {
        "lag1_log":la,"lag2_log":lb,"lag3_log":lc,
        "mean3_log":float(np.log1p(np.mean([a,b,c]))),
        "std3_log":float(np.std([la,lb,lc],ddof=0)),
        "growth_own":float(np.clip(la-lb,-2,2)),
        "target_month_sin":float(np.sin(2*np.pi*period/12)),
        "target_month_cos":float(np.cos(2*np.pi*period/12)),
        "peer_growth":float(peer[i,0]),"peer_density":float(peer[i,1]),
        "sku_index":float(i),
    }

def historical_graph(mat,t,history=6,k=5):
    """For each origin t, use only [t-history,t), hence no horizon leakage."""
    n=mat.shape[0]
    w=np.zeros((n,n),dtype=float)
    if t<history+1:return w,np.zeros((n,2))
    hist=np.log1p(np.maximum(mat[:,t-history:t],0))
    dlt=np.diff(hist,axis=1)
    for i in range(n):
        xi=dlt[i]
        if np.isfinite(xi).sum()<4:continue
        for j in range(i+1,n):
            xj=dlt[j]
            ok=np.isfinite(xi)&np.isfinite(xj)
            if ok.sum()<4:continue
            a,b=xi[ok],xj[ok]
            if np.std(a)<1e-4 or np.std(b)<1e-4:continue
            corr=float(np.corrcoef(a,b)[0,1])
            if not np.isfinite(corr):continue
            # 5 changes per 6 months are extremely fragile: shrink correlation.
            strength=max(0.,corr)*ok.sum()/(ok.sum()+8.)
            if strength > .075:w[i,j]=w[j,i]=strength
    if k<n:
        for i in range(n):
            good=np.flatnonzero(w[i]>0)
            if len(good)>k:
                w[i,good[np.argsort(w[i,good])[:-k]]]=0
    # Enforce symmetry only for mutual selected edges; edges are predictive not causal.
    w=np.minimum(w,w.T)
    peer=np.zeros((n,2))
    for i in range(n):
        neighbors=np.flatnonzero(w[i]>0)
        if not len(neighbors):continue
        ok=[j for j in neighbors if np.isfinite(mat[j,t-1]) and np.isfinite(mat[j,t-2])]
        if not ok:continue
        wg=w[i,ok]
        grows=np.log1p(np.maximum(mat[ok,t-1],0))-np.log1p(np.maximum(mat[ok,t-2],0))
        peer[i,0]=float(np.clip(np.sum(wg*grows)/wg.sum(),-2,2))
        peer[i,1]=len(ok)
    return w,peer

def make_dataset(mat,codes):
    rows=[]
    graphs={}
    for t in range(6,len(MN)):
        graph,peer=historical_graph(mat,t)
        graphs[t]=graph
        for i,code in enumerate(codes):
            for h in [1,2,3]:
                ti=t+h-1
                if ti>=len(MN):continue
                y=mat[i,ti]
                if not np.isfinite(y) or y<0:continue
                ft=numeric_features(mat,i,t,h,peer)
                if ft is None:continue
                last12=ti-12
                seasonal=mat[i,last12] if last12>=0 and last12<t and np.isfinite(mat[i,last12]) else np.nan
                row={
                    "sku_idx":i,"origin":MN[t],"target":MN[ti],"h":h,
                    "y":float(y),"Lag1":float(mat[i,t-1]),
                    "Mean3":float(np.mean(mat[i,t-3:t])),
                    "Seasonal12":float(seasonal) if np.isfinite(seasonal) else np.nan,
                    "PeerLag":float(max(0,mat[i,t-1]*np.exp(ft["peer_growth"]))),
                }
                row.update(ft)
                rows.append(row)
    return pd.DataFrame(rows),graphs

BASE=["lag1_log","lag2_log","lag3_log","mean3_log","std3_log",
      "growth_own","target_month_sin","target_month_cos"]
GRAPH=BASE+["peer_growth","peer_density"]
def fit_predict(ds,original_stats):
    preds=[]
    metrics=[]
    for h in (1,2,3):
        a=ds[ds.h==h].copy()
        train=a[a.target <= "202512"]
        val=a[a.target.between("202601","202602")]
        test=a[a.target.between("202603","202607")]
        assert train.target.max() <= "202512" and test.target.min()>="202603"
        if len(train)<75 or len(test)<30:
            metrics.append({"h":h,"status":"insufficient_train_or_test","train":len(train),"test":len(test)})
            continue
        # No data from test/validation for hyperparameter choice. Fixed preregistered settings.
        model_b=HistGradientBoostingRegressor(loss="squared_error",max_iter=85,max_leaf_nodes=12,
               min_samples_leaf=26,l2_regularization=8,learning_rate=.06,random_state=42)
        model_g=HistGradientBoostingRegressor(loss="squared_error",max_iter=85,max_leaf_nodes=12,
               min_samples_leaf=26,l2_regularization=8,learning_rate=.06,random_state=42)
        model_r=make_pipeline(StandardScaler(),Ridge(alpha=50))
        Y=np.log1p(np.maximum(0,train.y.to_numpy()))
        model_b.fit(train[BASE],Y)
        model_g.fit(train[GRAPH],Y)
        model_r.fit(train[BASE],Y)
        for method,estimated in [
            ("Lag1",test["Lag1"].to_numpy()),
            ("Mean3",test["Mean3"].to_numpy()),
            ("Seasonal12",test["Seasonal12"].to_numpy()),
            ("PeerLag",test["PeerLag"].to_numpy()),
            ("Ridge",np.expm1(model_r.predict(test[BASE]))),
            ("HistGB",np.expm1(model_b.predict(test[BASE]))),
            ("GraphHistGB",np.expm1(model_g.predict(test[GRAPH]))),
        ]:
            for (_,obs),v in zip(test.iterrows(),estimated):
                if not np.isfinite(v):continue
                preds.append({"h":h,"model":method,"sku_idx":int(obs["sku_idx"]),
                              "target":str(obs["target"]),"y":float(obs.y),
                              "yhat":max(0,float(v))})
        metrics.append({"h":h,"status":"computed","train":len(train),
                        "validation":len(val),"test":len(test),
                        "num_skus_train":int(train.sku_idx.nunique()),
                        "num_skus_test":int(test.sku_idx.nunique())})
    pred=pd.DataFrame(preds)
    return pred,metrics

def wmape(df):
    den=float(df.y.abs().sum())
    return float((df.y-df.yhat).abs().sum()/den) if den>0 else None
def summarize(pred):
    out=[]
    for (h,method),g in pred.groupby(["h","model"]):
        r={
           "h_months":int(h),"model":method,"n":len(g),
           "n_skus":int(g.sku_idx.nunique()),
           "test_months":int(g.target.nunique()),
           "mae":round(float(np.mean(np.abs(g.y-g.yhat))),6),
           "wmape":round(wmape(g),6) if wmape(g) is not None else None,
           "bias":round(float(np.mean(g.yhat-g.y)),6),
           "sum_y":round(float(g.y.sum()),4),
        }
        out.append(r)
    return out

def paired_ci(pred,a="HistGB",b="GraphHistGB",h=1,seed=42):
    d=pred[(pred.h==h)&pred.model.isin([a,b])]
    p=d.pivot_table(index=["target","sku_idx"],columns="model",values=["y","yhat"],aggfunc="first")
    try:
        da=p[[("y",a),("yhat",a),("yhat",b)]].dropna()
    except Exception:return None
    if da.empty:return None
    da.columns=["y","a","b"]
    dat=da.reset_index()
    monthly=np.sort(dat.target.unique())
    if len(monthly)<3:return None
    rng=np.random.default_rng(seed)
    vals=[]
    for _ in range(500):
        sampled=rng.choice(monthly,size=len(monthly),replace=True)
        g=pd.concat([dat.loc[dat.target==m] for m in sampled],ignore_index=True)
        den=g.y.abs().sum()
        if den>0:
            ea=(g.y-g.a).abs().sum()/den
            eb=(g.y-g.b).abs().sum()/den
            vals.append(ea-eb)
    if not vals:return None
    lo,hi=np.quantile(vals,[.025,.975])
    return {"comparison":f"{a}-minus-{b}","h":h,
            "matched_points":len(dat),"independent_month_blocks":len(monthly),
            "wmape_gain_absolute":float((dat.y-dat.a).abs().sum()/dat.y.abs().sum()-
                                         (dat.y-dat.b).abs().sum()/dat.y.abs().sum()),
            "boot95":[float(lo),float(hi)],
            "can_claim_statistical_advantage":bool(lo>0 or hi<0),
            "method":"500 month-block bootstrap, only five target months; CI exploratory"}

def charts(x,mat,codes,pred,graphs):
    def save(fig,name):
        fig.tight_layout(pad=1.15)
        fig.savefig(FIG/(name+".png"),dpi=230,bbox_inches="tight")
        fig.savefig(FIG/(name+".svg"),bbox_inches="tight")
        plt.close(fig)
    # An honest main figure: monthly observed volume & sampling coverage.
    month_sum=np.nansum(mat,axis=0)
    month_cov=np.sum(np.isfinite(mat),axis=0)
    fig,ax=plt.subplots(figsize=(10.4,3.5))
    ax.plot(range(len(MN)),month_sum,lw=2,color=COLORS["tree"],marker="o",markersize=3)
    ax.axvspan(11.5,13.5,color="#cf9e42",alpha=.11,label="Validation")
    ax.axvspan(13.5,18.5,color="#118c86",alpha=.08,label="Future test")
    ax.set_title("Observed product-month aggregate / recorded volume")
    ax.set_ylabel("Volume (source units; verify original unit)")
    ax.set_xlabel("2025-01 to 2026-07")
    ax.set_xticks(range(len(MN))[::2],labels=MN[::2],rotation=30)
    ax.grid(alpha=.14);ax.legend(frameon=False,fontsize=8)
    save(fig,"Fig01_monthly_actual_aggregate")
    fig,ax=plt.subplots(figsize=(10.4,3.4))
    ax.bar(range(len(MN)),month_cov,color=COLORS["tree"])
    ax.set_title("Available SKU count by observed month (missing is not zero)")
    ax.set_ylabel("SKUs with an observed volume");ax.set_xticks(range(len(MN))[::2],labels=MN[::2],rotation=30)
    save(fig,"Fig02_monthly_sku_coverage")
    # Missingness matrix ordered by product coverage.
    ordered=np.argsort(-np.sum(np.isfinite(mat),axis=1))
    fig,ax=plt.subplots(figsize=(10.4,5))
    ax.imshow(np.isfinite(mat[ordered]),aspect="auto",interpolation="nearest",cmap="Blues",vmin=0,vmax=1)
    ax.set_title("Product-month observed-data matrix (dark=recorded; light=missing)")
    ax.set_xlabel("Month");ax.set_ylabel("SKU sorted by coverage")
    ax.set_xticks(range(len(MN))[::2],labels=MN[::2],rotation=30)
    save(fig,"Fig03_recorded_coverage_matrix")
    if pred.empty:return
    main=pred[pred.h==1].copy()
    models=["Lag1","Mean3","Seasonal12","Ridge","HistGB","GraphHistGB"]
    # Paired evaluation only, for each model compare on available predictions.
    summary=summarize(pred)
    labels=sorted(set(q["h_months"] for q in summary))
    fig,ax=plt.subplots(figsize=(10,4.2))
    for m in models:
        dat=sorted((x for x in summary if x["model"]==m and x["wmape"] is not None),key=lambda x:x["h_months"])
        if dat:ax.plot([x["h_months"] for x in dat],[100*x["wmape"] for x in dat],marker="o",lw=2,label=m)
    ax.set_title("Out-of-time forecasting errors (actual monthly holdout)")
    ax.set_xlabel("Horizon (months)");ax.set_ylabel("wMAPE (%)")
    ax.set_xticks([1,2,3]);ax.grid(alpha=.15);ax.legend(ncol=3,fontsize=8,frameon=False)
    save(fig,"Fig04_multi_horizon_wmape")
    # Predefined SKU case selection by historical totals (not based on best predictions)
    historical=np.nansum(mat[:,:12],axis=1)
    tested=main.loc[main.model=="HistGB","sku_idx"].unique().astype(int)
    eligible=np.array([i for i in tested if np.isfinite(historical[i]) and historical[i]>0],dtype=int)
    if len(eligible):
        sortedids=eligible[np.argsort(historical[eligible])]
        cases=[sortedids[min(int((len(sortedids)-1)*q),len(sortedids)-1)] for q in (.25,.5,.9)]
        fig,axs=plt.subplots(3,1,figsize=(9.5,8),sharex=True)
        for ii,(ax,skuid) in enumerate(zip(axs,cases)):
            case=main[main.sku_idx==skuid]
            yy=case[case.model=="HistGB"].sort_values("target")
            gg=case[case.model=="GraphHistGB"].sort_values("target")
            if len(yy):
                xx=np.arange(len(yy))
                ax.plot(xx,yy.y,color="#161d2b",marker="o",lw=1.8,label="Observed")
                ax.plot(xx,yy.yhat,color=COLORS["tree"],marker="s",lw=1.5,label="HistGB")
                if len(gg)==len(yy):ax.plot(xx,gg.yhat,color=COLORS["graph"],marker="^",lw=1.5,label="GraphHistGB")
                ax.set_xticks(xx,labels=yy.target,rotation=25)
            ax.set_title(f"Train-volume quantile {int([.25,.5,.9][ii]*100)}: anonymized SKU",fontsize=10)
            ax.set_ylabel("Volume");ax.grid(alpha=.12)
        axs[0].legend(ncol=3,fontsize=8,frameon=False)
        save(fig,"Fig05_actual_vs_predicted_3_preselected_skus")
    # Graph visualization, one available pre-target date.
    n=mat.shape[0];graph=graphs.get(IDX["202601"])
    if graph is not None:
        totals=np.nansum(mat[:,:12],axis=1)
        ids=np.argsort(-totals)[:min(32,n)]
        sub=graph[np.ix_(ids,ids)]
        fig,ax=plt.subplots(figsize=(6.1,5.4))
        im=ax.imshow(sub,cmap="YlGnBu",vmin=0,vmax=max(0.2,float(np.nanmax(sub))))
        ax.set_title("Train-past-only graph weights, 32 leading SKUs")
        ax.set_xlabel("Anonymous SKU index");ax.set_ylabel("Anonymous SKU index")
        fig.colorbar(im,ax=ax,fraction=.046,pad=.04,label="Shrunk Pearson association")
        save(fig,"Fig06_graph_weights_training_only")
    fig,ax=plt.subplots(figsize=(9.3,3.8))
    by=[]
    for m in sorted(main.target.unique()):
        for model in ("HistGB","GraphHistGB"):
            d=main[(main.target==m)&(main.model==model)]
            v=wmape(d)
            by.append((m,model,v))
    for model,c in (("HistGB",COLORS["tree"]),("GraphHistGB",COLORS["graph"])):
        v=[x for x in by if x[1]==model]
        ax.plot([q[0] for q in v],[100*q[2] for q in v],marker="o",lw=1.9,label=model,color=c)
    ax.set_title("Monthly holdout wMAPE: no-graph vs graph features")
    ax.set_ylabel("wMAPE (%)");ax.set_xlabel("Independent test month")
    ax.legend(frameon=False);ax.grid(alpha=.15)
    save(fig,"Fig07_monthly_horizon1_graph_ablation")

def main():
    raw,orig,dups,bad=source_audit()
    x=load_panel()
    audit=verify_original(x,orig,dups,bad)
    codes,mat,cat,meta=prepare_matrix(x)
    audit.update({"periods":len(MN),"start":MN[0],"end":MN[-1],
                  "product_count":len(codes),"actual_month_sku_nonmissing":int(np.isfinite(mat).sum()),
                  "vol_missing_rows":int(x.vol.isna().sum()),
                  "vol_zero_rows":int((x.vol==0).sum()),
                  "vol_negative_rows":int((x.vol<0).sum()),
                  "coverage_per_month":{m:int(np.isfinite(mat[:,i]).sum()) for i,m in enumerate(MN)},
                  "unit":"source-defined volume; verify original business unit",
                  "daily_observations_verified":False,
                  "note":"The existing build_panel_from_raw_v17_3.py constructs one row per SKU per month; do not relabel as daily series."})
    output_json(RES/"01_data_audit.json",audit)
    ds,graphs=make_dataset(mat,codes)
    preds,run=fit_predict(ds,audit)
    if not preds.empty:
        metrics=summarize(preds)
        ci=paired_ci(preds)
        output_json(RES/"02_baseline_metrics.json",
                    {"experiment_provenance":"real GitHub repo XLS plus prior raw_panel_v1.csv",
                     "prediction_unit":"SKU x target month (not day)",
                     "models":["Lag1","Mean3","Seasonal12","PeerLag","Ridge","HistGB","GraphHistGB"],
                     "no_TFT_or_GAT_or_MTGNN_training":True,
                     "results":metrics,"grouping":run,"paired_month_bootstrap":ci,
                     "method_warning":"Only 12 training months, 2 validation and 5 future holdout: pilot evidence only."})
        preds.groupby(["h","target","model"]).agg(n=("y","size"),
                actual=("y","sum"),predicted=("yhat","sum")).reset_index().to_csv(
                        RES/"03_test_month_aggregates.csv",index=False,encoding="utf-8-sig")
    else:
        output_json(RES/"02_baseline_metrics.json",{"status":"insufficient sample","grouping":run})
    charts(x,mat,codes,preds,graphs)
    # Keep raw product identity, raw SKU predictions and source files out of paper output.
    print("REAL EXPERIMENT SUMMARY:",json.dumps({
       "rows":len(x),"skus":len(codes),"months":len(MN),
       "raw_match":audit.get("source_exact_fraction_of_matched"),
       "models_run":len(preds.model.unique()) if not preds.empty else 0,
       "output":str(RES.relative_to(ROOT))},ensure_ascii=False))
if __name__=="__main__":
    main()

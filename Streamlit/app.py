from pathlib import Path
import sys
import altair as alt
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Corps Recce Performance", page_icon="◈", layout="wide", initial_sidebar_state="expanded")
ROOT=Path(__file__).resolve().parents[1]; BOOK=ROOT/'Inputs'/'Step1_model_setup.xlsx'
FOLLOW_UP_ACTIONS=ROOT/'Inputs'/'follow_up_actions.csv'
if str(ROOT) not in sys.path: sys.path.insert(0,str(ROOT))
from DataEntry import app as data_entry
PAGES=['Home','Participants','Events','Progress','Follow-up','Data Entry']

def apply_kinpaku_theme():
    """Presentation-only Data Terminal styling for native Streamlit controls and views."""
    st.markdown("""
    <style>
    :root {
      --ink: #0c0b09;
      --deep: #080705;
      --raised: #15130f;
      --graphite: #201e19;
      --line: #2d2a23;
      --gold: #f1c95d;
      --gold-dim: #a98d43;
      --verdigris: #69c9bd;
      --text: #f4f1e8;
      --muted: #c8c2b6;
      --danger: #ed806d;
      --amber: #efbc5b;
    }
    .stApp { background: var(--ink); color: var(--text); }
    .stApp, .stApp * { font-family: "Avenir Next", "Segoe UI", sans-serif; }
    .material-symbols-rounded, .material-symbols-outlined, [class*="material-symbols"] { font-family: "Material Symbols Rounded" !important; }
    [data-testid="stAppViewContainer"] { background: var(--ink); }
    [data-testid="stHeader"] { background: transparent; }
    [data-testid="stSidebar"] { background: #0a0907; border-right: 1px solid var(--line); }
    [data-testid="stExpandSidebarButton"], [data-testid="stSidebarCollapseButton"] { display: none; }
    [data-testid="stSidebar"] > div:first-child { padding: 1.35rem .85rem; }
    [data-testid="stSidebar"] h2 { color: var(--text); letter-spacing: -.015em; font-size: 1.05rem; margin: .3rem .5rem .2rem; }
    [data-testid="stSidebar"] h2 + div p { color: var(--muted); font-size: .78rem; margin: 0 .5rem 1.7rem; }
    [data-testid="stSidebar"] [data-testid="stCaptionContainer"] p { color: var(--muted); font-size: .76rem; font-weight: 500; margin-top: 1.7rem; }
    .block-container { max-width: 1560px; padding: 1.85rem 2rem 3.5rem; }
    h1 { color: var(--text) !important; font-size: clamp(1.65rem, 2.4vw, 2.25rem) !important; letter-spacing: -.025em; font-weight: 600 !important; line-height: 1.15 !important; margin: .2rem 0 .35rem !important; }
    h2, h3 { color: var(--text) !important; letter-spacing: -.01em; font-weight: 600 !important; }
    h3 { margin-top: 2rem !important; font-size: 1rem !important; }
    [data-testid="stCaptionContainer"] p { color: var(--muted); font-size: .82rem; }
    [data-testid="stMetric"] { background: transparent; border: 0; border-left: 1px solid var(--line); border-radius: 0; padding: .35rem 1rem .45rem; min-height: 78px; }
    [data-testid="stColumn"]:first-child [data-testid="stMetric"] { border-left: 0; padding-left: 0; }
    [data-testid="stMetricLabel"] { color: var(--muted) !important; font-size: .78rem; letter-spacing: 0; text-transform: none; }
    [data-testid="stMetricValue"] { color: var(--text) !important; font-variant-numeric: tabular-nums; font-size: 1.7rem; letter-spacing: -.025em; }
    [data-testid="stMetricDelta"] { color: var(--verdigris) !important; }
    [data-testid="stVerticalBlockBorderWrapper"] { border-color: var(--line) !important; border-radius: 6px; }
    [data-testid="stAlert"] { background: #221e13; border: 0; border-left: 2px solid var(--amber); color: var(--text); border-radius: 0; }
    [data-testid="stDataFrame"], [data-testid="stTable"] { border: 0; border-radius: 0; overflow: hidden; }
    [data-testid="stDataFrame"] * { font-variant-numeric: tabular-nums; }
    [data-baseweb="select"] > div, [data-baseweb="input"] > div { background: var(--raised) !important; border-color: var(--line) !important; color: var(--text) !important; border-radius: 5px !important; }
    [data-baseweb="input"] input, [data-baseweb="select"] input { color: var(--text) !important; }
    [data-baseweb="popover"] { background: var(--raised) !important; color: var(--text) !important; }
    [data-baseweb="radio"] label { color: var(--muted) !important; }
    [data-baseweb="radio"] input:checked + div { background-color: var(--gold) !important; }
    [data-testid="stSidebar"] [data-testid="stButton"] button { width: 100%; min-height: 2.45rem; justify-content: flex-start; padding-left: .85rem; background: transparent; color: var(--muted); border: 0; border-radius: 4px; font-size: .9rem; font-weight: 500; }
    [data-testid="stSidebar"] [data-testid="stButton"] button:hover { background: var(--graphite); color: var(--text); }
    [data-testid="stSidebar"] [data-testid="stButton"] button[kind="primary"] { background: #16130e; color: var(--text); border-left: 1px solid var(--gold); border-radius: 0 4px 4px 0; }
    [data-testid="stButton"] button { background: transparent; color: var(--text); border: 1px solid var(--line); border-radius: 5px; transition: background .16s ease, color .16s ease, transform .12s ease; }
    [data-testid="stButton"] button:hover { background: var(--graphite); color: var(--text); transform: translateY(-1px); }
    [data-testid="stButton"] button[kind="primary"] { background: var(--text); border-color: var(--text); color: var(--deep); font-weight: 600; }
    [data-testid="stButton"] button[kind="primary"]:hover { background: #ffffff; color: var(--deep); }
    button:focus-visible, input:focus-visible { outline: 2px solid var(--gold) !important; outline-offset: 2px !important; }
    hr { border-color: var(--line) !important; }
    [data-testid="stSidebar"] hr { margin: 1.2rem 0 !important; }
    ::selection { background: var(--gold); color: #17130a; }
    ::-webkit-scrollbar { width: 10px; height: 10px; }
    ::-webkit-scrollbar-track { background: var(--deep); }
    ::-webkit-scrollbar-thumb { background: #4a4335; border: 2px solid var(--deep); border-radius: 6px; }
    ::-webkit-scrollbar-thumb:hover { background: var(--gold-dim); }
    @media (max-width: 900px) {
      .block-container { padding: 1.2rem 1rem 3rem; }
      h1 { font-size: 1.75rem !important; }
      [data-testid="stMetric"] { min-height: 72px; padding: .5rem .7rem; }
    }
    </style>
    """, unsafe_allow_html=True)

def load():
    def read(n,h): return pd.read_excel(BOOK,sheet_name=n,header=h).dropna(how='all')
    ev,sub,cy,pa=read('DimEvent',1),read('DimSubEvent',1),read('DimCycle',1),read('DimParticipant',1)
    # Workbook keeps human-readable notes beneath reference tables; only numeric keys are model records.
    ev=ev[pd.to_numeric(ev.EventID,errors='coerce').notna()].copy(); sub=sub[pd.to_numeric(sub.SubEventID,errors='coerce').notna()].copy(); cy=cy[pd.to_numeric(cy.CycleID,errors='coerce').notna()].copy()
    pa=pa[pd.to_numeric(pa.ParticipantID,errors='coerce').notna()].copy()
    fa=read('FactResult',0); fa=fa[fa.ResultID.astype(str).str.match(r'^R\d+$',na=False)].copy()
    fa['DateAttempted']=pd.to_datetime(fa.DateAttempted,errors='coerce')
    for c in ['FinalMarks','MinMarks','MaxMarks','RawPerformance']: fa[c]=pd.to_numeric(fa[c],errors='coerce')
    d=(fa.merge(pa,on='ParticipantID',how='left',validate='many_to_one',suffixes=('','_p'))
        .merge(sub,on='SubEventID',how='left',validate='many_to_one',suffixes=('','_sub'))
        .merge(ev,on='EventID',how='left',validate='many_to_one',suffixes=('','_event'))
        .merge(cy,on='CycleID',how='left',validate='many_to_one'))
    d['Completed']=d.Status.eq('Completed'); d['ScorePct']=(d.FinalMarks/d.MaxMarks).where(d.Completed & d.MaxMarks.gt(0)); d['Target']=d.FinalMarks.ge(d.MinMarks)
    return d,ev,sub,cy,pa

def score_view(d,method):
    c=d[d.Completed].sort_values(['ParticipantID','SubEventID','DateAttempted','ResultID'])
    if method=='Latest': return c.groupby(['ParticipantID','SubEventID'],as_index=False).tail(1)
    if method=='Best': return c.sort_values('FinalMarks').groupby(['ParticipantID','SubEventID'],as_index=False).tail(1)
    latest_cycle=c.CycleID.max(); return c[c.CycleID.eq(latest_cycle)] # Official: completed scores in most recent cycle only

def event_scores(d,method):
    s=score_view(d,method).copy()
    if s.empty: return s
    s['WeightedScore']=s.ScorePct*s.Weight; s['WeightedMinimum']=(s.MinMarks/s.MaxMarks)*s.Weight
    group=['ParticipantID','ParticipantName','EventID','EventName']
    out=s.groupby(group,as_index=False).agg(ScorePct=('WeightedScore','sum'),MinimumPct=('WeightedMinimum','sum'),DateAttempted=('DateAttempted','max'),CompletedSubEvents=('SubEventID','nunique'),ResultID=('ResultID','last'),CycleID=('CycleID','max'),CycleName=('CycleName','last'))
    out['FinalMarks']=out.ScorePct.mul(100); out['MinMarks']=out.MinimumPct.mul(100); out['MaxMarks']=100; out['Target']=out.ScorePct.ge(out.MinimumPct); out['Completed']=True; out['Status']='Completed'
    return out

def sub_event_summary(d,method):
    s=score_view(d,method)
    return s.groupby(['EventID','EventName','SubEventID','SubEventName'],as_index=False).agg(Candidates=('ParticipantID','nunique'),AverageScore=('ScorePct','mean'),TargetRate=('Target','mean'),Weight=('Weight','first'))

@st.cache_data
def load_follow_up_actions():
    if not FOLLOW_UP_ACTIONS.exists(): return pd.DataFrame(columns=['ResultID','Recommended action'])
    actions=pd.read_csv(FOLLOW_UP_ACTIONS,dtype={'ResultID':str,'Recommended action':str})
    return actions[['ResultID','Recommended action']].drop_duplicates('ResultID',keep='last')

def save_follow_up_actions(actions):
    actions=actions[['ResultID','Recommended action']].drop_duplicates('ResultID',keep='last').copy()
    actions['Recommended action']=actions['Recommended action'].fillna('').str.strip().replace('', 'Review / schedule reassessment')
    existing=load_follow_up_actions()
    saved=pd.concat([existing[~existing.ResultID.isin(actions.ResultID)],actions],ignore_index=True)
    saved.to_csv(FOLLOW_UP_ACTIONS,index=False)
    load_follow_up_actions.clear()

def go(page,participant=None,event=None):
    st.session_state.page=page
    if participant is not None: st.session_state.participant=participant
    if event is not None: st.session_state.event=event
    st.rerun()

def nav_button(label,page):
    active=st.session_state.get('page','Home')==page
    if st.sidebar.button(label,width='stretch',key='nav_'+page,type='primary' if active else 'secondary'): go(page)

def reset_filters():
    for key in ['cycles','statuses','score_range']:
        st.session_state.pop(key,None)

def filters(d):
    st.sidebar.caption('FILTERS')
    fields=[('CycleName','Test','cycles')]
    for col,label,key in fields:
        vals=sorted(d[col].dropna().unique()); chosen=st.sidebar.multiselect(label,vals,key=key)
        if chosen:d=d[d[col].isin(chosen)]
    statuses=st.sidebar.multiselect('Assessment status',sorted(d.Status.dropna().unique()),key='statuses')
    if statuses:d=d[d.Status.isin(statuses)]
    score_range=st.sidebar.slider('Completed score range (%)',0,100,(0,100),key='score_range')
    if score_range!=(0,100):d=d[d.ScorePct.mul(100).between(*score_range)]
    method=st.sidebar.radio('Score selection',['Latest','Best','Official'],horizontal=True,key='method')
    st.sidebar.button('↺ Reset filters',width='stretch',on_click=reset_filters)
    return d,method

def header(title,crumb):
    st.title(title); st.caption(f'Home > {crumb}  ·  Last refreshed: {pd.Timestamp.now().strftime("%d %b %Y %H:%M")}')
    if st.session_state.get('page','Home')!='Home':
        a,b=st.columns([1,8]);
        if a.button('⌂ Home'):go('Home');
        if b.button('← Back'):go('Home')

def metrics(d,s):
    complete=d[d.Completed]; attention=(~s.Target).sum()
    cols=st.columns(5)
    vals=[('Overall score',f'{s.ScorePct.mean()*100:.1f}%'),('Participants',str(d.ParticipantID.nunique())),('Events',str(d.EventID.nunique())),('Completion',f'{len(complete)/len(d):.0%}' if len(d) else '—'),('Needs attention',str(attention))]
    for c,(n,v) in zip(cols,vals):c.metric(n,v)

def home(d,method):
    header('Recce Performance','Home — Management Overview'); s=event_scores(d,method); metrics(d,s)
    x,y=st.columns([2,1]);
    with x:
        st.subheader('Overall performance across cycles')
        trend=pd.concat([event_scores(cycle,'Latest') for _,cycle in d.groupby('CycleID')]).groupby('CycleName',as_index=False).ScorePct.mean(); st.line_chart(trend.set_index('CycleName'))
    with y:
        st.subheader('Assessment completion'); st.bar_chart(d.Status.value_counts())
    x,y=st.columns([2,1]);
    with x:
        st.subheader('Average score by event')
        avg=s.groupby('EventName').ScorePct.mean().sort_values(); st.bar_chart(avg)
    with y:
        st.subheader('Urgent follow-up')
        risk=s[~s.Target][['ParticipantName','EventName','ScorePct']].head(6); st.dataframe(risk,hide_index=True,width='stretch')
    st.subheader('Participant × event — latest completed score')
    heat=event_scores(d,'Latest').pivot_table(index='ParticipantName',columns='EventName',values='ScorePct',aggfunc='first'); st.dataframe(heat.style.format('{:.0%}'),width='stretch')
    c=st.columns(3)
    for col,label,page in zip(c,['View all participants','View all events','View follow-up'],['Participants','Events','Follow-up']):
        if col.button(label,width='stretch'):go(page)

def participants(d,method):
    header('Participants','Participants — Directory'); s=event_scores(d,method)
    q=st.text_input('Search participant or ID').lower(); p=s.groupby(['ParticipantID','ParticipantName'],dropna=False).agg(EventsCompleted=('EventID','nunique'),OverallScore=('ScorePct','mean'),TargetStatus=('Target','mean'),LastAssessment=('DateAttempted','max')).reset_index()
    if q:p=p[p.ParticipantName.str.lower().str.contains(q)|p.ParticipantID.astype(str).str.contains(q)]
    selection=st.dataframe(p,hide_index=True,width='stretch',on_select='rerun',selection_mode='single-row',key='participant_directory')
    if selection.selection.rows: go('Progress',participant=p.iloc[selection.selection.rows[0]].ParticipantName)
    choice=st.selectbox('Open participant profile',sorted(p.ParticipantName.unique()) if len(p) else [])
    if st.button('View participant profile',disabled=not choice):go('Progress',participant=choice)

def events(d,ev,sub,method):
    header('Events','Events — Catalogue'); s=event_scores(d,method)
    x=ev.copy(); a=s.groupby('EventID').agg(Participants=('ParticipantID','nunique'),AverageScore=('ScorePct','mean')).reset_index(); component_counts=sub.groupby('EventID').SubEventID.nunique().rename('Markable components').reset_index(); x=x.merge(a,on='EventID',how='left').merge(component_counts,on='EventID',how='left'); st.dataframe(x,hide_index=True,width='stretch')
    st.subheader('Sub-event score overview')
    components=sub_event_summary(d,method); st.bar_chart(components.sort_values('AverageScore').set_index('SubEventName').AverageScore); st.dataframe(components[['EventName','SubEventName','Weight','Candidates','AverageScore','TargetRate']],hide_index=True,width='stretch')
    choice=st.selectbox('Open event analysis',sorted(x.EventName.unique()))
    if st.button('View event analysis'):go('Event Analysis',event=choice)

def event_analysis(d,method):
    header('Event Analysis','Events — Event Analysis'); choice=st.selectbox('Event',sorted(d.EventName.dropna().unique()),index=max(0,sorted(d.EventName.dropna().unique()).index(st.session_state.get('event',sorted(d.EventName.dropna().unique())[0])))); event_data=d[d.EventName.eq(choice)]; s=event_scores(event_data,method)
    a,b,c,d1=st.columns(4); a.metric('Average',f'{s.ScorePct.mean():.0%}');b.metric('Median',f'{s.ScorePct.median():.0%}');c.metric('Completion',f'{s.ParticipantID.nunique()/max(1,event_data.ParticipantID.nunique()):.0%}');d1.metric('Marked components',event_data.SubEventID.nunique())
    rank=s.sort_values('ScorePct',ascending=False); rank['Rank']=rank.ScorePct.rank(method='min',ascending=False).astype(int); st.bar_chart(rank.set_index('ParticipantName').ScorePct); st.dataframe(rank[['Rank','ParticipantName','ScorePct','Status']],hide_index=True,width='stretch')
    st.subheader('Sub-event breakdown')
    components=sub_event_summary(event_data,method).sort_values('AverageScore'); st.bar_chart(components.set_index('SubEventName').AverageScore); st.dataframe(components[['SubEventName','Weight','Candidates','AverageScore','TargetRate']],hide_index=True,width='stretch')
    person=st.selectbox('Open selected participant',sorted(rank.ParticipantName.unique()))
    if st.button('View participant-event detail'):go('Progress',participant=person,event=choice)

def progress(d,method):
    header('Participant Profile & Progress','Participants > Participant Profile'); people=sorted(d.ParticipantName.dropna().unique()); person=st.selectbox('Participant',people,index=max(0,people.index(st.session_state.get('participant',people[0])))); p=d[d.ParticipantName.eq(person)]; participant_events=sorted(p.EventName.dropna().unique()); selected_event=st.session_state.get('event'); options=['All events']+participant_events; event=st.selectbox('View results for',options,index=options.index(selected_event) if selected_event in participant_events else 0); s=score_view(p,method)
    s=event_scores(p,method)
    if event=='All events':
        metrics(p,s); st.subheader('All event results'); st.bar_chart(s.set_index('EventName').ScorePct); st.dataframe(s[['EventName','FinalMarks','MinMarks','MaxMarks','ScorePct','Target','DateAttempted','CompletedSubEvents']],hide_index=True,width='stretch')
        return
    h=p[p.EventName.eq(event)].sort_values('DateAttempted'); history=h[h.Completed & h.DateAttempted.notna() & h.FinalMarks.notna()]; selected=event_scores(h,method)
    metrics(h,selected); st.subheader(f'{event} — performance history')
    if history.empty: st.info('No completed attempts are available for this event.')
    else:
        dated=history.assign(WeightedScore=history.ScorePct*history.Weight).groupby('DateAttempted').WeightedScore.sum()
        st.line_chart(dated)
        if history.SubEventID.nunique()>1:
            st.subheader('Sub-event score history')
            components=history.groupby(['DateAttempted','SubEventName'],as_index=False).ScorePct.mean()
            chart=alt.Chart(components).mark_line(point=True).encode(
                x=alt.X('DateAttempted:T',title='Test date'),
                y=alt.Y('ScorePct:Q',title='Score',scale=alt.Scale(domain=[0,1]),axis=alt.Axis(format='%')),
                color=alt.Color('SubEventName:N',legend=None,scale=alt.Scale(range=['#69c9bd','#f1c95d','#efbc5b','#ed806d','#b0ebe3'])),
                tooltip=[alt.Tooltip('DateAttempted:T',title='Test date'),alt.Tooltip('SubEventName:N',title='Sub-event'),alt.Tooltip('ScorePct:Q',title='Score',format='.1%')]
            ).properties(height=280)
            st.altair_chart(chart,width='stretch')
    first=history.sort_values('DateAttempted').groupby('DateAttempted').apply(lambda x:(x.ScorePct*x.Weight).sum()).iloc[0] if not history.empty else None; latest=history.sort_values('DateAttempted').groupby('DateAttempted').apply(lambda x:(x.ScorePct*x.Weight).sum()).iloc[-1] if not history.empty else None
    st.write(f'Plain-language interpretation: {"Insufficient completed history." if first is None or latest is None else f"weighted score changed by {(latest-first)*100:.0f} percentage points across {h.Completed.sum()} completed component attempts."}')
    st.subheader('Test dates and attempt results')
    st.dataframe(h[['DateAttempted','SubEventName','Weight','RawPerformance','MinMarks','MaxMarks','FinalMarks','Status','Comments']],hide_index=True,width='stretch')
    if st.button('View follow-up actions'):go('Follow-up',participant=person,event=event)

def followup(d,method):
    header('Follow-up','Follow-up — Action Centre'); s=score_view(d,method); r=s[~s.Target].copy(); r['Reason']=r.apply(lambda x:'Below target' if pd.notna(x.ScorePct) else 'Missing assessment',axis=1); r=r.merge(load_follow_up_actions(),on='ResultID',how='left'); r['Recommended action']=r['Recommended action'].fillna('Review / schedule reassessment'); r['Status']='Not started';r['Due date']='Needs supervisor confirmation'; columns=['ResultID','ParticipantName','EventName','Reason','Recommended action','Due date','Status']; edited=st.data_editor(r[columns],hide_index=True,width='stretch',disabled=[column for column in columns if column!='Recommended action'],key='follow_up_actions')
    if st.button('Save follow-up actions'): save_follow_up_actions(edited[['ResultID','Recommended action']]); st.success('Follow-up actions saved.')

def main():
    apply_kinpaku_theme()
    d,ev,sub,cy,pa=load(); st.sidebar.markdown('## Corps Recce\nPerformance workspace');
    for page in PAGES:nav_button(page,page)
    page=st.session_state.get('page','Home'); st.sidebar.divider();d,method=filters(d)
    if page=='Home':home(d,method)
    elif page=='Participants':participants(d,method)
    elif page=='Events':events(d,ev,sub,method)
    elif page=='Event Analysis':event_analysis(d,method)
    elif page=='Progress':progress(d,method)
    elif page=='Follow-up':followup(d,method)
    elif page=='Data Entry':data_entry.render()
main()

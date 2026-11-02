from datetime import date
from pathlib import Path
import os
import tempfile

import openpyxl
import pandas as pd
import streamlit as st

ROOT=Path(__file__).resolve().parents[1]
BOOK=ROOT/'Inputs'/'Step1_model_setup.xlsx'
STATUSES=['Completed','Absent','Exempt','Disqualified','Not Attempted']

def header_row(ws,key):
    for row in range(1,ws.max_row+1):
        if key in [cell.value for cell in ws[row]]: return row
    raise ValueError(f'{key} header was not found in {ws.title}.')

def table_rows(ws,key):
    row=header_row(ws,key); headers=[cell.value for cell in ws[row]]
    return headers,[dict(zip(headers,values)) for values in ws.iter_rows(min_row=row+1,values_only=True) if values[0] is not None]

def load_model(book=BOOK):
    workbook=openpyxl.load_workbook(book,data_only=True)
    def frame(name,key):
        headers,records=table_rows(workbook[name],key)
        return pd.DataFrame(records,columns=headers)
    return frame('DimParticipant','ParticipantID'),frame('DimEvent','EventID'),frame('DimSubEvent','SubEventID'),frame('DimCycle','CycleID'),frame('FactResult','ResultID')

def next_identifier(values,prefix,width=4):
    numbers=[int(str(value)[len(prefix):]) for value in values if str(value).startswith(prefix) and str(value)[len(prefix):].isdigit()]
    return f'{prefix}{max(numbers,default=0)+1:0{width}d}'

def validate_result(record,sub_event):
    if record['Status'] not in STATUSES: raise ValueError('Select a valid assessment status.')
    if not record['Assessor'].strip(): raise ValueError('Assessor is required.')
    if not record['DateAttempted']: raise ValueError('Date attempted is required.')
    if record['Status']=='Completed':
        if record['RawPerformance'] is None: raise ValueError('Completed results require raw performance.')
        if not 0<=record['FinalMarks']<=sub_event['MaxMarks']: raise ValueError('Final marks must be within the configured range.')
    else:
        record['RawPerformance']=None; record['FinalMarks']=0
    record['MinMarks']=sub_event['MinMarks']; record['MaxMarks']=sub_event['MaxMarks']; record['RawMeasure']=sub_event['UnitOfMeasure']; record['HigherIsBetter']=sub_event['HigherIsBetter']
    return record

def save_workbook(workbook,book=BOOK):
    with tempfile.NamedTemporaryFile(dir=book.parent,suffix='.xlsx',delete=False) as temporary:
        temporary_path=Path(temporary.name)
    try:
        workbook.save(temporary_path); os.replace(temporary_path,book)
    finally:
        temporary_path.unlink(missing_ok=True)

def append_result(record,book=BOOK):
    workbook=openpyxl.load_workbook(book); ws=workbook['FactResult']; headers,_=table_rows(ws,'ResultID')
    record['ResultID']=next_identifier([cell.value for cell in ws['A'][1:]],'R'); record['AttemptID']=next_identifier([row[1].value for row in ws.iter_rows(min_row=2)],'M')
    ws.append([record.get(column) for column in headers]); save_workbook(workbook,book)
    return record['ResultID']

def update_result(result_id,record,book=BOOK):
    workbook=openpyxl.load_workbook(book); ws=workbook['FactResult']; headers,_=table_rows(ws,'ResultID'); result_column=headers.index('ResultID')+1
    for row in range(2,ws.max_row+1):
        if ws.cell(row,result_column).value==result_id:
            for column,value in record.items():
                if column in headers: ws.cell(row,headers.index(column)+1).value=value
            save_workbook(workbook,book); return
    raise ValueError(f'Result {result_id} was not found.')

def append_reference(sheet,key,record,book=BOOK):
    workbook=openpyxl.load_workbook(book); ws=workbook[sheet]; headers,_=table_rows(ws,key); ws.append([record.get(column) for column in headers]); save_workbook(workbook,book)

def result_form(prefix,participants,events,sub_events,cycles,existing=None):
    participant_options={str(row.ParticipantName):int(row.ParticipantID) for _,row in participants.iterrows()}
    event_options={row.EventName:int(row.EventID) for _,row in events.iterrows()}
    participant_values=list(participant_options.values()); event_values=list(event_options.values())
    existing_sub=sub_events[sub_events.SubEventID.eq(existing['SubEventID'])].iloc[0] if existing else None
    participant_label=st.selectbox('Participant',participant_options,index=participant_values.index(int(existing['ParticipantID'])) if existing else 0,key=f'{prefix}_participant')
    parent_column,sub_column=st.columns(2)
    with parent_column: event_name=st.selectbox('Parent event',event_options,index=event_values.index(int(existing_sub.EventID)) if existing else 0,key=f'{prefix}_event')
    event_id=event_options[event_name]
    available=sub_events[sub_events.EventID.eq(event_id)]
    sub_options={row.SubEventName:int(row.SubEventID) for _,row in available.iterrows()}
    sub_key=f'{prefix}_sub'
    if st.session_state.get(sub_key) not in sub_options: st.session_state.pop(sub_key,None)
    if len(available)>1:
        with sub_column: sub_name=st.selectbox('Markable sub-event',sub_options,key=sub_key)
        sub_id=sub_options[sub_name]
    else:
        with sub_column: st.caption('This assessment is marked as one overall result.')
        sub_id=int(available.iloc[0].SubEventID)
    sub=available[available.SubEventID.eq(sub_id)].iloc[0].to_dict()
    cycle_options={row.CycleName:int(row.CycleID) for _,row in cycles.iterrows()}
    cycle_values=list(cycle_options.values())
    cycle_name=st.selectbox('Assessment cycle',cycle_options,index=cycle_values.index(int(existing['CycleID'])) if existing else 0,key=f'{prefix}_cycle')
    status=st.selectbox('Status',STATUSES,index=STATUSES.index(existing['Status']) if existing else 0,key=f'{prefix}_status')
    attempted=st.date_input('Date attempted',value=pd.to_datetime(existing['DateAttempted']).date() if existing else date.today(),key=f'{prefix}_date')
    completed=status=='Completed'; default_marks=float(existing['FinalMarks']) if existing and completed else float(sub['MinMarks'])
    st.caption(f"Valid marks: 0–{sub['MaxMarks']}; target: {sub['MinMarks']}; weight: {sub['Weight']:.0%}")
    final_marks=st.number_input('Final marks',min_value=0.0,max_value=float(sub['MaxMarks']),value=default_marks,step=1.0,disabled=not completed,key=f'{prefix}_marks')
    raw=st.number_input('Raw performance',min_value=0.0,value=float(existing['RawPerformance']) if existing and completed and pd.notna(existing['RawPerformance']) else default_marks,step=1.0,disabled=not completed,key=f'{prefix}_raw')
    assessor=st.text_input('Assessor',value=str(existing['Assessor']) if existing else '',key=f'{prefix}_assessor')
    comments=st.text_area('Comments',value=str(existing['Comments']) if existing else '',key=f'{prefix}_comments')
    submitted=st.button('Save result',key=f'{prefix}_save')
    record={'ParticipantID':participant_options[participant_label],'SubEventID':sub_id,'CycleID':cycle_options[cycle_name],'RawPerformance':raw if completed else None,'FinalMarks':final_marks if completed else 0,'Status':status,'Assessor':assessor,'Comments':comments,'DateAttempted':attempted.isoformat()}
    return submitted,record,sub

def amend_form(existing,sub):
    status=st.selectbox('Status',STATUSES,index=STATUSES.index(existing['Status']),key='amend_status')
    attempted=st.date_input('Date attempted',value=pd.to_datetime(existing['DateAttempted']).date(),key='amend_date')
    completed=status=='Completed'; default_marks=float(existing['FinalMarks']) if completed else float(sub['MinMarks'])
    st.caption(f"Valid marks: 0–{sub['MaxMarks']}; target: {sub['MinMarks']}; weight: {sub['Weight']:.0%}")
    final_marks=st.number_input('Final marks',min_value=0.0,max_value=float(sub['MaxMarks']),value=default_marks,step=1.0,disabled=not completed,key='amend_marks')
    raw=st.number_input('Raw performance',min_value=0.0,value=float(existing['RawPerformance']) if completed and pd.notna(existing['RawPerformance']) else default_marks,step=1.0,disabled=not completed,key='amend_raw')
    assessor=st.text_input('Assessor',value=str(existing['Assessor']),key='amend_assessor')
    comments=st.text_area('Comments',value=str(existing['Comments']),key='amend_comments')
    submitted=st.button('Save changes',key='amend_save')
    return submitted,{'ParticipantID':existing['ParticipantID'],'SubEventID':existing['SubEventID'],'CycleID':existing['CycleID'],'RawPerformance':raw if completed else None,'FinalMarks':final_marks if completed else 0,'Status':status,'Assessor':assessor,'Comments':comments,'DateAttempted':attempted.isoformat()}

def go(page):
    st.session_state.entry_page=page; st.rerun()

def back_to_home():
    if st.button('← Back to home'): go('Home')

def render():
    st.title('Recce Performance Data Entry')
    st.caption('Enter and correct assessment records in the structured Excel workbook.')
    if not BOOK.exists(): st.error(f'Workbook not found: {BOOK}'); st.stop()
    participants,events,sub_events,cycles,facts=load_model(); page=st.session_state.get('entry_page','Home')
    if page=='Home':
        st.subheader('What would you like to do?')
        one,two,three=st.columns(3)
        if one.button('Add new participant or test',width='stretch'): go('Reference')
        if two.button('New assessment',width='stretch'): go('New assessment')
        if three.button('Amend assessment',width='stretch'): go('Amend assessment')
        return
    back_to_home()
    if page=='New assessment':
        st.subheader('New assessment')
        submitted,record,sub=result_form('new_result',participants,events,sub_events,cycles)
        if submitted:
            try:
                result_id=append_result(validate_result(record,sub)); st.success(f'{result_id} saved to FactResult.'); st.rerun()
            except ValueError as error: st.error(str(error))
            except PermissionError: st.error('Close the Excel workbook, then save again.')
            except Exception: st.error('Oops, we ran into an error. Check the fields and try again.')
    elif page=='Amend assessment':
        st.subheader('Amend assessment')
        try:
            test_options={row.CycleName:int(row.CycleID) for _,row in cycles.iterrows()}
            test_name=st.selectbox('Test',test_options,key='amend_test'); test_id=test_options[test_name]
            test_facts=facts[facts.CycleID.eq(test_id)]
            candidate_rows=participants[participants.ParticipantID.isin(test_facts.ParticipantID)]
            candidate_options={str(row.ParticipantName):int(row.ParticipantID) for _,row in candidate_rows.iterrows()}
            candidate_name=st.selectbox('Candidate',candidate_options,key='amend_candidate'); candidate_id=candidate_options[candidate_name]
            candidate_facts=test_facts[test_facts.ParticipantID.eq(candidate_id)].merge(sub_events[['SubEventID','EventID','SubEventName']],on='SubEventID').merge(events,on='EventID')
            assessment_options={row.EventName:int(row.EventID) for _,row in candidate_facts.drop_duplicates('EventID').iterrows()}
            assessment_name=st.selectbox('Available assessment',assessment_options,key='amend_event'); assessment_id=assessment_options[assessment_name]
            assessment_facts=candidate_facts[candidate_facts.EventID.eq(assessment_id)]
            if assessment_facts.SubEventID.nunique()>1:
                component_options={row.SubEventName:int(row.SubEventID) for _,row in assessment_facts.drop_duplicates('SubEventID').iterrows()}
                component_name=st.selectbox('Markable sub-event',component_options,key='amend_component'); assessment_facts=assessment_facts[assessment_facts.SubEventID.eq(component_options[component_name])]
            if len(assessment_facts)>1:
                attempt_options={f"{row.ResultID} · {pd.to_datetime(row.DateAttempted).date()}":row.ResultID for _,row in assessment_facts.iterrows()}
                attempt_name=st.selectbox('Attempt',attempt_options,key='amend_attempt'); result_id=attempt_options[attempt_name]
            else: result_id=assessment_facts.iloc[0].ResultID
            existing=facts[facts.ResultID.eq(result_id)].iloc[0].to_dict(); sub=sub_events[sub_events.SubEventID.eq(existing['SubEventID'])].iloc[0].to_dict()
            submitted,record=amend_form(existing,sub)
            if submitted:
                try:
                    update_result(result_id,validate_result(record,sub)); st.success(f'{result_id} updated.'); st.rerun()
                except ValueError as error: st.error(str(error))
                except PermissionError: st.error('Close the Excel workbook, then save again.')
                except Exception: st.error('Oops, we ran into an error. Check the fields and try again.')
        except (IndexError,KeyError,TypeError,ValueError):
            st.error('Oops, we ran into an error. Check the selected result and required fields, then try again.')
    else:
        st.subheader('Add new participant or test')
        left,right=st.columns(2)
        with left:
            st.markdown('#### Add participant')
            with st.form('add_participant'):
                name=st.text_input('Participant name'); submitted=st.form_submit_button('Add participant')
            if submitted:
                if not name.strip(): st.error('Participant name is required.')
                elif name.strip().lower() in participants.ParticipantName.astype(str).str.lower().tolist(): st.error('A participant with this name already exists.')
                else:
                    participant_id=int(participants.ParticipantID.max())+1
                    try:
                        append_reference('DimParticipant','ParticipantID',{'ParticipantID':participant_id,'ParticipantName':name.strip(),'PersonCode':f'P{participant_id:03d}','Status':'Active'}); st.success('Participant added.'); st.rerun()
                    except PermissionError: st.error('Close the Excel workbook, then save again.')
                    except Exception: st.error('Oops, we ran into an error. Check the fields and try again.')
        with right:
            st.markdown('#### Add assessment test')
            with st.form('add_cycle'):
                name=st.text_input('Test name'); start=st.date_input('Start date'); end=st.date_input('End date'); description=st.text_input('Description'); cycle_type=st.selectbox('Test type',['Baseline','Review','Follow-up']); submitted=st.form_submit_button('Add test')
            if submitted:
                if not name.strip() or end<start: st.error('Enter a test name and an end date on or after the start date.')
                elif name in cycles.CycleName.astype(str).tolist(): st.error('Test name must be unique.')
                else:
                    try:
                        append_reference('DimCycle','CycleID',{'CycleID':int(cycles.CycleID.max())+1,'CycleName':name.strip(),'StartDate':start,'EndDate':end,'Description':description.strip(),'CycleType':cycle_type}); st.success('Test added.'); st.rerun()
                    except PermissionError: st.error('Close the Excel workbook, then save again.')
                    except Exception: st.error('Oops, we ran into an error. Check the fields and try again.')
        st.subheader('Protected scoring structure')
        st.dataframe(sub_events.merge(events,on='EventID')[['EventName','SubEventName','Weight','MinMarks','MaxMarks','UnitOfMeasure']],hide_index=True,width='stretch')

def main():
    st.set_page_config(page_title='Recce Performance Data Entry', page_icon='✎', layout='wide')
    render()

if __name__=='__main__': main()

from pathlib import Path
import shutil
import tempfile
from streamlit.testing.v1 import AppTest

source=Path('DataEntry/app.py').read_text()
namespace={'__file__':str(Path('DataEntry/app.py').resolve()),'__name__':'DataEntry.app'}
exec(source,namespace)
for label in ['Add new participant or test','New assessment','Amend assessment','← Back to home']:
    assert label in source

book=Path(tempfile.mkdtemp())/'model.xlsx'
shutil.copy(Path('Inputs/Step1_model_setup.xlsx'),book)
participants,events,subs,cycles,facts=namespace['load_model'](book)
sub=subs.iloc[0].to_dict()
record={'ParticipantID':int(participants.iloc[0].ParticipantID),'SubEventID':int(sub['SubEventID']),'CycleID':int(cycles.iloc[0].CycleID),'RawPerformance':75,'FinalMarks':75,'Status':'Completed','Assessor':'Test assessor','Comments':'Test entry','DateAttempted':'2026-05-01'}
result_id=namespace['append_result'](namespace['validate_result'](record,sub),book)
assert result_id not in facts.ResultID.tolist()
namespace['update_result'](result_id,{'Comments':'Corrected entry'},book)
facts=namespace['load_model'](book)[4]
assert facts.iloc[-1].ResultID==result_id and facts.iloc[-1].Comments=='Corrected entry'
app=AppTest.from_file('app.py'); app.run(timeout=30)
assert {'Add new participant or test','New assessment','Amend assessment'} <= {button.label for button in app.button}
app.button[[button.label for button in app.button].index('New assessment')].click().run(timeout=30)
parent=app.selectbox[[field.label for field in app.selectbox].index('Parent event')]
parent.set_value('12 Km').run(timeout=30)
assert 'Markable sub-event' not in [field.label for field in app.selectbox]
parent.set_value('SA firing').run(timeout=30)
assert 'Markable sub-event' in [field.label for field in app.selectbox]
app=AppTest.from_file('app.py'); app.run(timeout=30)
app.button[[button.label for button in app.button].index('Amend assessment')].click().run(timeout=30)
assert not app.exception, app.exception
assert {'Test','Candidate','Available assessment'} <= {field.label for field in app.selectbox}
assert 'Person code' not in [field.label for field in app.text_input]
print('PASS: assessment entry and correction write valid workbook records.')

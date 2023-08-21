import datetime
from pathlib import Path
from docxtpl import DocxTemplate

base_dir=Path(__file__).parent
word_template = base_dir / "demo_offer_format.docx"

today=datetime.datetime.today()

doc=DocxTemplate(word_template)
context={
    "DATE":today.strftime("%Y-%m-%d"),
    "NAME":"abc",
    "POSITION":"PRO Team member"
}
doc.render(context)
doc.save(base_dir / "offerlettertest1.docx")
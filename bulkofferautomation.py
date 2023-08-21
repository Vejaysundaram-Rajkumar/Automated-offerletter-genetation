import datetime
from pathlib import Path
from docxtpl import DocxTemplate
import pandas as pd


base_dir=Path(__file__).parent
word_template = base_dir / "demo_offer_format.docx"
excel_file=base_dir / "demo_list.xlsx"
output_folder=base_dir / "Offerletters"

output_folder.mkdir(exist_ok=True)


df=pd.read_excel(excel_file,sheet_name="Sheet1")

df["TODAY"] = pd.to_datetime(df["TODAY"]).dt.time

for i in df.to_dict(orient="records"):
    doc=DocxTemplate(word_template)
    doc.render(i)
    out_path=output_folder / f"{i['NAME']}-offerletter.docx"
    doc.save(out_path)



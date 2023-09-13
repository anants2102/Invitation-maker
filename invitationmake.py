from pathlib import Path
import pandas as pd
from docxtpl import DocxTemplate 
import docx2pdf as dd

base_dir = Path(__file__).parent
invitation_template = base_dir / "invitation.docx"
excel = base_dir / "Details.xlsx"
output_dir = base_dir / "built_invitations"
output_dir.mkdir(exist_ok=True)
df = pd.read_excel(excel, sheet_name= "Sheet1")

for record in df.to_dict(orient = "records"):
     doc = DocxTemplate(invitation_template )
     doc.render(record)
     output_path = output_dir / f"{record['Company']} Invitation MMMUT.docx"
     doc.save(output_path)

out = base_dir/ "buitl_invitations_pdf"
out.mkdir(exist_ok=True)
dd.convert("built_invitations/", out)
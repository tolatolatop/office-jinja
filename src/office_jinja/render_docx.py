import os
from docxtpl import DocxTemplate
from .models import Context, Template


def render_all_docs_from_template(template: Template) -> None:
    for template_file in template.templates:
        if not template_file.endswith(".docx"):
            continue
        render_docx(template_file, template.context)


def render_docx(template_file: str, context: Context) -> None:
    doc = DocxTemplate(template_file)
    doc.render(context.data)
    for pic_file, pic_path in context.pic.items():
        doc.replace_pic(pic_file, pic_path)
    doc.save(template_file.replace(".docx", ".rendered.docx"))

import os
import logging
from docxtpl import DocxTemplate
from .models import Context, Template


def render_all_docs_from_template(template: Template) -> None:
    """
    渲染所有docx模板
    """
    for template_file in template.templates:
        if not template_file.endswith(".docx"):
            continue
        render_docx(template_file, template.context)


def render_docx(template_file: str, context: Context) -> None:
    """
    渲染单个docx模板
    """
    doc = DocxTemplate(template_file)
    doc.render(context.data)
    replace_all_pictures(doc, context)
    doc.save(template_file.replace(".docx", ".rendered.docx"))


def replace_all_pictures(doc: DocxTemplate, context: Context) -> None:
    """
    替换所有图片
    """
    for pic_file, pic_path in context.pic.items():
        if not os.path.exists(pic_path):
            logging.error(f"Picture file {pic_path} not found")
            raise FileNotFoundError(f"Picture file {pic_path} not found")
        if pic_file not in doc.get_pic_map():
            logging.debug(f"Picture file {pic_file} not found")
            continue
        doc.replace_pic(pic_file, pic_path)

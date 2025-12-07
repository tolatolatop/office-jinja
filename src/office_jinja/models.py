import os
from ruamel.yaml import YAML
from pydantic import BaseModel, field_validator, Field


class TemplateFile(BaseModel):
    path: str
    output_path: str


def get_output_path(path: str) -> str:
    return path.replace(".docx", ".rendered.docx").replace(".xlsx", ".rendered.xlsx")


class TemplateCollection(BaseModel):
    templates: list[TemplateFile] = []

    @field_validator('templates', mode="before")
    def validate_templates(cls, v: list[TemplateFile | str]) -> list[TemplateFile]:
        templates = []
        for template in v:
            if isinstance(template, str):
                tmpl = TemplateFile(
                    path=template,
                    output_path=get_output_path(template)
                )
                templates.append(tmpl)
            else:
                templates.append(template)
        return templates


class Context(BaseModel):
    data: dict
    pic: dict | None = None


class Template(TemplateCollection):
    name: str
    context: Context

    @classmethod
    def model_validate_yaml(cls, file: str) -> "Template":
        with open(file, encoding="utf-8") as f:
            yaml = YAML()
            yaml.indent(mapping=2, sequence=4, offset=2)
            return cls.model_validate(yaml.load(f))

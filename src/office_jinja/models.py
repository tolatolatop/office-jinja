from pydantic import BaseModel


class Template(BaseModel):
    name: str
    templates: list[str]
    context: dict

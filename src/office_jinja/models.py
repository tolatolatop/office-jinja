from ruamel.yaml import YAML
from pydantic import BaseModel


class Context(BaseModel):
    data: dict
    pic: dict


class Template(BaseModel):
    name: str
    templates: list[str]
    context: Context

    @classmethod
    def model_validate_yaml(cls, file: str) -> "Template":
        with open(file, encoding="utf-8") as f:
            yaml = YAML()
            yaml.indent(mapping=2, sequence=4, offset=2)
            return cls.model_validate(yaml.load(f))

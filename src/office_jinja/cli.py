import yaml
from click import argument, command, version_option

from office_jinja.models import Template

version = "0.1.0"


@command()
@version_option(version=version)
@argument("config")
def generate(config: str):
    with open(config, encoding="utf-8") as f:
        data = yaml.safe_load(f)
    template = Template.model_validate(data)
    print(template)


if __name__ == "__main__":
    generate()

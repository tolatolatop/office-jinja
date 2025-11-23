import yaml
from click import argument, command, version_option

from office_jinja.models import Template

version = "0.1.0"


@command()
@version_option(version=version)
@argument("config")
def generate(config: str):
    template = Template.model_validate_yaml(config)
    print(template)


if __name__ == "__main__":
    generate()

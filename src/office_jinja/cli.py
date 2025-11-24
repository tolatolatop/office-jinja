import yaml
from click import argument, command, version_option

from office_jinja.models import Template
from office_jinja.render_docx import render_all_docs_from_template

version = "0.1.0"


@command()
@version_option(version=version)
@argument("config")
def generate(config: str):
    template = Template.model_validate_yaml(config)
    print(template)
    render_all_docs_from_template(template)


if __name__ == "__main__":
    generate()

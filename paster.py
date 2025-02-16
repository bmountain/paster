from src.paster_utils import get_config, get_children
from src.paster_app import ExcelPaster


config = get_config()
children = get_children(config.dirname)
paster = ExcelPaster(config, children)
paster.run()

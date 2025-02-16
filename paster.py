from src.paster_utils import get_params, get_children
from src.paster_app import ExcelPaster


params = get_params()
children = get_children(params.dirname)
paster = ExcelPaster(params, children)
paster.run()

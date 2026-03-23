"""
UI Windows package
Contiene le finestre principali dell'applicazione.
"""

def get_main_module():
    """
    Dinamically imports the main DataFlow module.
    This is needed because the main file has spaces in its name.
    """
    import importlib.util
    import os
    
    # Find the main file
    current_dir = os.path.dirname(__file__)
    workspace_dir = os.path.dirname(os.path.dirname(current_dir))
    main_file = os.path.join(workspace_dir, 'dataflow.py')
    
    if not os.path.exists(main_file):
        raise ImportError(f"Main file not found: {main_file}")
    
    # Load the module
    spec = importlib.util.spec_from_file_location("dataflow_main", main_file)
    dataflow_main = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(dataflow_main)
    
    return dataflow_main

from .view_request_window import ViewRequestWindow
from .edit_suppliers_window import EditSuppliersWindow
from .edit_reference_window import EditReferenceWindow
from .notes_window import NotesWindow
from .purchase_order_window import PurchaseOrderWindow
from .attachment_window import AttachmentWindow
from .sqdc_analysis_window import SQDCAnalysisWindow

# MainWindow non ancora estratta - rimane in dataflow.py
# from .main_window import MainWindow

__all__ = [
    'ViewRequestWindow',
    'EditSuppliersWindow',
    'EditReferenceWindow',
    'NotesWindow',
    'PurchaseOrderWindow',
    'AttachmentWindow',
    'SQDCAnalysisWindow',
    'get_main_module'
]

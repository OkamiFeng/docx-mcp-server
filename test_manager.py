from docx_manager import DocxManager
import os

def test_manager():
    mgr = DocxManager()
    print("Testing creation...")
    print(mgr.create_new())
    
    print("Testing add heading...")
    print(mgr.add_heading("Hello World Docx", level=1))
    
    print("Testing add paragraph...")
    print(mgr.add_paragraph("This is a test paragraph.", style="Normal"))
    
    print("Testing add table...")
    print(mgr.add_table(2, 2, [["A1", "B1"], ["A2", "B2"]]))
    
    save_path = os.path.abspath("test_output.docx")
    print(f"Testing save to {save_path}...")
    print(mgr.save_document(save_path))
    
    print("Testing read structure...")
    print(mgr.get_structure())
    
    print("DONE")

if __name__ == "__main__":
    test_manager()

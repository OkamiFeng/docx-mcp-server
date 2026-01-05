from docx_manager import DocxManager
import os
import glob
import shutil

def run_integration_test():
    """
    Simulates the AI Agent workflow:
    1. Read source docx from 'test/'
    2. Extract content and images
    3. 'Summarize' (mocked)
    4. Create new docx with summary and images
    """
    print("--- Starting Integration Test ---")
    
    # Setup paths
    test_dir = os.path.abspath("test")
    output_dir = os.path.abspath("test_output_artifacts")
    if os.path.exists(output_dir):
        shutil.rmtree(output_dir)
    os.makedirs(output_dir)
    
    # 1. Find source file
    docx_files = glob.glob(os.path.join(test_dir, "*.docx"))
    if not docx_files:
        print("No .docx files found in test directory.")
        return

    source_file = docx_files[0]
    print(f"Processing source file: {source_file}")
    
    # 2. Initialize Manager
    mgr = DocxManager()
    
    # 3. Load and Read
    print(mgr.load_document(source_file))
    structure = mgr.read_full_content()
    print(f"Source Stats: {structure['paragraphs'][:3]}... (Truncated)")
    print(f"Found {structure['images_count']} images.")
    
    # 4. Extract Images
    images_dir = os.path.join(output_dir, "extracted_images")
    extracted_images = mgr.extract_images(images_dir)
    print(f"Extracted images to: {images_dir}")
    print(f"Image files: {[os.path.basename(p) for p in extracted_images]}")
    
    # 5. Mock Summary Generation
    summary_text = f"Summary of {os.path.basename(source_file)}:\n"
    summary_text += f"This document contains {len(structure['paragraphs'])} paragraphs and {structure['tables_count']} tables.\n"
    summary_text += "Key points extracted from reading the original text..."
    
    # 6. Create New Document
    print("Creating new summary document...")
    mgr.create_new()
    mgr.add_heading("Automated Summary Report", level=1)
    mgr.add_paragraph(summary_text)
    
    if extracted_images:
        mgr.add_heading("Extracted Visuals", level=2)
        mgr.add_paragraph("Below are the images found in the original document:")
        for img_path in extracted_images:
            try:
                mgr.add_paragraph(f"Image: {os.path.basename(img_path)}")
                mgr.add_image(img_path, width_inches=4.0)
                print(f"Inserted image: {os.path.basename(img_path)}")
            except Exception as e:
                print(f"Failed to insert {img_path}: {e}")
    
    # 7. Save
    output_docx = os.path.join(output_dir, "Summary_Report.docx")
    mgr.save_document(output_docx)
    print(f"Generated summary document saved to: {output_docx}")
    print("--- Integration Test Complete ---")

if __name__ == "__main__":
    run_integration_test()

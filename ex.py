import win32com.client
import os
import sys

def get_acad_app():
    """
    Try to connect to an existing AutoCAD.Application.
    Falls back to GetObject if Dispatch fails.
    """
    try:
        return win32com.client.Dispatch("AutoCAD.Application")
    except Exception:
        try:
            return win32com.client.GetObject(None, "AutoCAD.Application")
        except Exception as e:
            print("❌ Could not connect to AutoCAD:", e)
            sys.exit(1)

def get_filename_from_doc(doc):
    """
    Return the file name for a given document.
    If doc.FullName is empty (unsaved), use doc.Name.
    """
    full = getattr(doc, "FullName", "")
    if full and full.strip():
        return os.path.basename(full)
    else:
        return getattr(doc, "Name", "<unnamed>")

def main():
    acad = get_acad_app()
    
    # Active document
    try:
        active = acad.ActiveDocument
        print("Active Drawing:", get_filename_from_doc(active))
    except Exception as e:
        print("⚠️  No active document or failed to retrieve it:", e)
    
    # All open documents by direct iteration
    docs = acad.Documents
    try:
        iterator = list(docs)  # force COM enumeration into a Python list
    except Exception:
        iterator = None

    if iterator:
        print("\nAll Open Drawings:")
        for idx, doc in enumerate(iterator, start=1):
            try:
                name = get_filename_from_doc(doc)
                print(f"  {idx}. {name}")
            except Exception as e:
                print(f"  {idx}. <failed to read document>: {e}")
    else:
        # fallback: maybe no docs or enumeration failed
        print("ℹ️  No documents are currently open or enumeration failed.")

if __name__ == "__main__":
    main()

import streamlit.web.cli as stcli
import os, sys

def resolve_path(path):
    if getattr(sys, "frozen", False):
        basedir = sys._MEIPASS
    else:
        basedir = os.path.dirname(__file__)
    return os.path.join(basedir, path)

if __name__ == "__main__":
    # Path to the streamlit app script
    app_path = resolve_path("streamlit_app.py")
    
    # Set up arguments for streamlit
    sys.argv = [
        "streamlit",
        "run",
        app_path,
        "--global.developmentMode=false",
    ]
    
    # Run streamlit
    sys.exit(stcli.main())

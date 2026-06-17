import os
import sys
import streamlit.web.cli as stcli

if __name__ == '__main__':
    # Buscamos la ruta absoluta de tu script principal
    script_path = os.path.join(os.path.dirname(__file__), 'app_v10.py')
    
    # Simulamos el comando 'streamlit run app.py' desde código
    sys.argv = ["streamlit", "run", script_path, "--global.developmentMode=false"]
    sys.exit(stcli.main())
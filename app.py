from flask import Flask, request, send_file, render_template
import pandas as pd

app = Flask(__name__)

# Ruta para cargar y procesar los archivos
@app.route('/process', methods=['POST'])
def process_files():
    # Cargar el archivo .xlsx enviado
    input_file = request.files['input_file']
    template_file_path = 'PlantillaSTEP4.xlsx'
    
    # Leer el archivo de entrada
    df_input = pd.read_excel(input_file)

    # Cargar la plantilla
    df_template = pd.read_excel(template_file_path)

    # Rellenar la plantilla según las condiciones
    for index, row in df_input.iterrows():
        # Asignar valores de "Name", "Surname" y "E-mail"
        df_template.at[index + 6, 'Name'] = row['Name']  # C se llena con A
        df_template.at[index + 6, 'Surname'] = row['Surname']  # D se llena con B
        df_template.at[index + 6, 'Primary email'] = row['E-mail']  # E se llena con O

        # Asignar "Primary phone"
        pcc_value = row['Va a ser PCC?']
        market_value = row['Market']
        
        if pcc_value == 'Y':
            if market_value == 'DACH':
                df_template.at[index + 6, 'Primary phone'] = '/+4940210918145 /+43122709858 /+41445295828'
                df_template.at[index + 6, 'Workgroup'] = 'D_PCC'
                df_template.at[index + 6, 'Team'] = 'Team_D_CCH_PCC_1'
                df_template.at[index + 6, 'Is PCC'] = 'Y'
                df_template.at[index + 6, 'Campaign Level'] = 'Agent'
            elif market_value == 'France':
                df_template.at[index + 6, 'Primary phone'] = '/+33180037979'
                df_template.at[index + 6, 'Workgroup'] = 'F_PCC'
                df_template.at[index + 6, 'Team'] = 'Team_F_CCH_PCC_1'
                df_template.at[index + 6, 'Is PCC'] = 'Y'
                df_template.at[index + 6, 'Campaign Level'] = 'Agent'
            elif market_value == 'Spain':
                df_template.at[index + 6, 'Primary phone'] = '/+34932952130'
                df_template.at[index + 6, 'Workgroup'] = 'E_PCC'
                df_template.at[index + 6, 'Team'] = 'Team_E_CCH_PCC_1'
                df_template.at[index + 6, 'Is PCC'] = 'Y'
                df_template.at[index + 6, 'Campaign Level'] = 'Agent'
            elif market_value == 'Italy':
                df_template.at[index + 6, 'Primary phone'] = '/+390109997099'
                df_template.at[index + 6, 'Workgroup'] = 'I_PCC'
                df_template.at[index + 6, 'Team'] = 'Team_I_CCH_PCC_1'
                df_template.at[index + 6, 'Is PCC'] = 'Y'
                df_template.at[index + 6, 'Campaign Level'] = 'Agent'
        elif pcc_value == 'N':
            if market_value == 'DACH':
                df_template.at[index + 6, 'Workgroup'] = 'D_Outbound'
                df_template.at[index + 6, 'Team'] = 'Team_D_CCH_B2C_1'
                df_template.at[index + 6, 'Is PCC'] = 'N'
                df_template.at[index + 6, 'Campaign Level'] = 'Agent'
            elif market_value == 'France':
                df_template.at[index + 6, 'Workgroup'] = 'F_Outbound'
                df_template.at[index + 6, 'Team'] = 'Team_F_CCH_B2C_1'
                df_template.at[index + 6, 'Is PCC'] = 'N'
                df_template.at[index + 6, 'Campaign Level'] = 'Agent'
            elif market_value == 'Spain':
                df_template.at[index + 6, 'Workgroup'] = 'E_Outbound'
                df_template.at[index + 6, 'Team'] = 'Team_E_CCH_B2C_1'
                df_template.at[index + 6, 'Is PCC'] = 'N'
                df_template.at[index + 6, 'Campaign Level'] = 'Agent'
        elif pcc_value == 'TL':
            if market_value == 'DACH':
                df_template.at[index + 6, 'Workgroup'] = 'D_PCC'
                df_template.at[index + 6, 'Team'] = 'Team_D_CCH_PCC_1'
                df_template.at[index + 6, 'Is PCC'] = 'N'
                df_template.at[index + 6, 'Campaign Level'] = 'Team Leader'
            elif market_value == 'France':
                df_template.at[index + 6, 'Workgroup'] = 'F_PCC'
                df_template.at[index + 6, 'Team'] = 'Team_F_CCH_PCC_1'
                df_template.at[index + 6, 'Is PCC'] = 'N'
                df_template.at[index + 6, 'Campaign Level'] = 'Team Leader'
            elif market_value == 'Spain':
                df_template.at[index + 6, 'Workgroup'] = 'E_PCC'
                df_template.at[index + 6, 'Team'] = 'Team_E_CCH_PCC_1'
                df_template.at[index + 6, 'Is PCC'] = 'N'
                df_template.at[index + 6, 'Campaign Level'] = 'Team Leader'
            elif market_value == 'Italy':
                df_template.at[index + 6, 'Workgroup'] = 'I_PCC'
                df_template.at[index + 6, 'Team'] = 'Team_I_CCH_PCC_1'
                df_template.at[index + 6, 'Is PCC'] = 'N'
                df_template.at[index + 6, 'Campaign Level'] = 'Team Leader'
        elif pcc_value == 'DS':
            if market_value in ['DACH', 'France', 'Spain']:
                df_template.at[index + 6, 'Workgroup'] = 'D_Outbound'
                df_template.at[index + 6, 'Team'] = 'Team_D_CCH_B2C_1'
                df_template.at[index + 6, 'Is PCC'] = 'N'
                df_template.at[index + 6, 'Campaign Level'] = 'Agent'

        # Rellenar otros campos
        df_template.at[index + 6, 'CTI User'] = row['B2E User Name']
        df_template.at[index + 6, 'TTG UserID 1'] = row['B2E User Name']
    
    # Guardar la plantilla rellenada
    output_file_path = 'Rellenada_PlantillaSTEP4.xlsx'
    df_template.to_excel(output_file_path, index=False)

    # Enviar el archivo generado
    return send_file(output_file_path, as_attachment=True)

# Ruta para mostrar el formulario de carga
@app.route('/')
def upload_form():
    return render_template('upload.html')

if __name__ == '__main__':
    app.run(debug=True)

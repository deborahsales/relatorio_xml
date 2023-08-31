import zipfile
import os
import xml.etree.ElementTree as ET
from openpyxl import Workbook
from datetime import datetime

def alunos(zip_file_path, nome_arquivo):

    # Cria a pasta de destino, se ela não existir
    pasta_destino = "./arquivos_extraidos"
    if not os.path.exists(pasta_destino):
        os.makedirs(pasta_destino)
        
    # Extrair os arquivos do ZIP
    with zipfile.ZipFile(zip_file_path, 'r') as zip_ref:
        # Encontre e extraia o arquivo de atividades
        activity_xml_file_path = None
        user_xml_file_path = None
        for file_name in zip_ref.namelist():
            if file_name.startswith("activity_"):
                activity_xml_file_path = file_name
            elif file_name.startswith("users_"):
                user_xml_file_path = file_name

            if activity_xml_file_path and user_xml_file_path:
                break

        if not activity_xml_file_path or not user_xml_file_path:
            print("Arquivos de atividades e/ou usuários não encontrados no ZIP.")
            return

        zip_ref.extract(activity_xml_file_path, "./arquivos_extraidos")
        zip_ref.extract(user_xml_file_path, "./arquivos_extraidos")
    
    activity_xml_path = os.path.join("./arquivos_extraidos", activity_xml_file_path)
    user_xml_path = os.path.join("./arquivos_extraidos", user_xml_file_path)

    # Processar o arquivo de usuários para criar um dicionário com os dados
    user_data = {}
    user_tree = ET.parse(user_xml_path)
    user_root = user_tree.getroot()
    for user in user_root.findall(".//user"):
        user_id = user.find('userId').text
        first_name = user.find('firstName').text
        last_name = user.find('lastName').text
        screen_name = user.find('screenName').text
        user_data[user_id] = {'firstName': first_name, 'lastName': last_name, 'screenName': screen_name}

    # Criar um novo arquivo Excel
    workbook = Workbook()
    sheet = workbook.active

    # Encontrar o nome da licença
    license_element = user_root.find('.//license')
    license_name = license_element.text.strip()
    
    # Insira as fórmulas nas células
    sheet['B1'] = f'Relatório de Uso - {license_name}'
    sheet['D1'] = 'Professor criando 2023:'
    sheet['G1'] = 'Alunos usando 2023:'

    header = ['ID atividade', 'Atividade', 'ID Professor', 'Professor', 'Data criação', 'ID aluno', 'Usuario', 'Data realização', 'Nota', 'Aprovado']
    sheet.append(header)

    # Processar os blocos de atividades
    activity_tree = ET.parse(activity_xml_path)
    activity_root = activity_tree.getroot()

    unique_user_ids = set()
    unique_alunos_ids = set()

    for activity in activity_root.findall(".//activity"):
        act_id = activity.find('actId').text
        title = activity.find('title').text
        create_date_text = activity.find('createDate').text
        create_date = datetime.strptime(create_date_text, '%Y-%m-%d %H:%M:%S').date()
        user_id_prof = activity.find('userId').text
        prof_info = user_data.get(user_id_prof, {})
        user_prof = prof_info.get('screenName', '')

        if create_date.year == 2023:  # Verifica se o ano é 2023
            unique_user_ids.add(user_id_prof)  # Adiciona user_id_prof ao conjunto

        activity_results = activity.findall('.//activityResult')
        for result in activity_results:
            user_id = result.find('userId').text
            result_value = result.find('result').text
            if result.find('passed').text == 'true':
                passed = "Sim" 
            else:
                passed = "Não"
            start_date_text = result.find('startDate').text
            start_date = datetime.strptime(start_date_text, '%Y-%m-%d %H:%M:%S').date()

            user_info = user_data.get(user_id, {})
            screen_name = user_info.get('screenName', '')

            sheet.append([act_id, title, user_id_prof, user_prof, create_date, user_id, screen_name, start_date, result_value, passed])

            if start_date.year == 2023:  # Verifica se o ano é 2023
                unique_alunos_ids.add(user_id)  # Adiciona user_id_prof ao conjunto

    # Congelar a primeira linha (cabeçalho)
    sheet.freeze_panes = 'A3'

    # Ajustar a largura das colunas e definir formatos
    for idx, column in enumerate(sheet.columns, start=1):
        max_length = 0
        column = [cell for cell in column]
        for cell in column:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(cell.value)
            except:
                pass
        adjusted_width = min(max_length + 2, 30)  # Definir a largura máxima como 200
        sheet.column_dimensions[column[0].column_letter].width = adjusted_width
        
        # Definir o formato de data para coluna E (createDate) e H (startDate)
        if idx == 5:  # Coluna E (createDate)
            for cell in column:
                cell.number_format = 'YYYY-MM-DD'  # Definir o formato de data
                
        if idx == 8:  # Coluna H (startDate)
            for cell in column:
                cell.number_format = 'YYYY-MM-DD'  # Definir o formato de data

    
    sheet['D1'] = f'Professor criando 2023: {len(unique_user_ids)}'
    sheet['G1'] = f'Alunos usando 2023: {len(unique_alunos_ids)}'

    # Salvar a planilha com o nome da licença
    excel_file_path = f'{nome_arquivo}_activity_{license_name}.xlsx'
    workbook.save(excel_file_path)

    print(f'Planilha "{excel_file_path}" criada com sucesso!')
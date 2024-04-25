import pandas as pd
import googlemaps
import numpy as np

gmaps = googlemaps.Client(key='KEY_PRIMO_API')

file_path = r'CAMINHO_DA_PLANILHA'

enderecos1 = pd.read_excel(file_path, sheet_name='Enderecos1')
enderecos2 = pd.read_excel(file_path, sheet_name='Enderecos2')

print("Colunas em Endereços1:", enderecos1.columns)
print("Colunas em Endereços2:", enderecos2.columns)

enderecos1['Mais Próximo'] = ''
enderecos1['Endereço Mais Próximo'] = ''  # Novo campo para o endereço
enderecos1['Distância (m)'] = np.inf

for index1, row1 in enderecos1.iterrows():
    loc1 = (row1['Latitude'], row1['Longitude'])
    nome_proximo = None
    endereco_proximo = None  # Variável para armazenar o endereço
    menor_distancia = np.inf
    print(f"Processando {row1['Nome do contato']}...")  # Print de progresso

    for index2, row2 in enderecos2.iterrows():
        loc2 = (row2['Latitude'], row2['Longitude'])

        distancia = gmaps.distance_matrix(loc1, loc2, mode='driving')['rows'][0]['elements'][0]
        dist_metros = distancia['distance']['value'] if 'distance' in distancia else np.inf

        if dist_metros < menor_distancia:
            menor_distancia = dist_metros
            nome_proximo = row2['Nome do contato']
            endereco_proximo = row2['Endereço']  # Captura o endereço

    enderecos1.at[index1, 'Mais Próximo'] = nome_proximo
    enderecos1.at[index1, 'Endereço Mais Próximo'] = endereco_proximo  # Salva o endereço
    enderecos1.at[index1, 'Distância (m)'] = menor_distancia

updated_file_path = r'LOCAL_PRA_SALVAR'
enderecos1.to_excel(updated_file_path, index=False)

print("Processamento concluído. Os resultados foram salvos no arquivo 'comparar_atualizado.xlsx'.")

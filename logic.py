import re
import asyncio
import streamlit as st
import pandas as pd

# Define o caminho de rede
CAMINHO_REDE = "\\\\10.111.200.59\\Controle\\Estoque"

# Verifica se a mensagem contém um serial válido
def verificSerial(mensagem):
    padrao = r'\b[A-Z0-9]{6,}\b'  # Ajuste o padrão conforme necessário
    match = re.search(padrao, mensagem)
    if match:
        return match.group()
    else:
        return None

# Consulta a planilha local em busca do número de série
def consultDados(Serial):
    try:
        df = pd.read_excel(f"{CAMINHO_REDE}\\Relatório Personalizado 14-05-2024.xlsx")
        if Serial:
            resultado = df[df['NÚMERO DO SERIAL'] == Serial]
            if not resultado.empty:
                # Guarda os valores da planilha de consulta
                NomeDispositivo = resultado['NOME DO DISPOSITIVO'].iloc[0]
                ModeloDispositivo = resultado['APELIDO'].iloc[0]
                SerialDispositivo = resultado['NÚMERO DO SERIAL'].iloc[0]
                ProcessadorUsado = resultado['PROCESSADOR'].iloc[0]
                MemoriaTotal = resultado['MEMÓRIA RAM TOTAL'].iloc[0]
                ArmazenamentoInterno = resultado['ARMAZENAMENTO INTERNO TOTAL'].iloc[0]
                ObservacaoDispositivo = resultado['OBSERVAÇÃO DO DISPOSITIVO'].iloc[0]
                return resultado, NomeDispositivo, ModeloDispositivo, SerialDispositivo, ProcessadorUsado, MemoriaTotal, ArmazenamentoInterno, ObservacaoDispositivo
        return "Nada encontrado.", None, None, None, None, None, None, None
    except Exception as e:
        return f"Erro ao consultar a planilha de estoque: {e}", None, None, None, None, None, None, None

# Adiciona uma nova linha à planilha de estoque com os valores guardados
def adicDados(NomeDispositivo, ModeloDispositivo, SerialDispositivo, ProcessadorUsado, MemoriaTotal, ArmazenamentoInterno, ObservacaoDispositivo, TipoDispositivo): 
    try:
        nova_linha = pd.DataFrame({
            'NomeDispositivo': [NomeDispositivo],
            'ModeloDispositivo': [ModeloDispositivo],
            'SerialDispositivo': [SerialDispositivo],
            'ProcessadorUsado': [ProcessadorUsado],
            'MemoriaTotal': [MemoriaTotal],
            'ArmazenamentoInterno': [ArmazenamentoInterno],
            'ObservacaoDispositivo': [ObservacaoDispositivo],
            'TipoDispositivo': [TipoDispositivo]  # Nova coluna para armazenar o tipo do dispositivo
        })
        
        df = pd.read_excel(f"{CAMINHO_REDE}\\Estoque atual.xlsx")
        df = pd.concat([df, nova_linha], ignore_index=True)
        df.to_excel(f"{CAMINHO_REDE}\\Estoque atual.xlsx", index=False)
        return "Boa bixo! Equipamento adicionado no estoque."
    except Exception as e:
        return f"Erro ao adicionar nova linha à planilha de estoque: {e}"

# Exclui uma linha da planilha de estoque com base no tipo de dispositivo e ID ou Serial
def exclDados(mensagem, serial_dispositivo):
    try:
        # Padroes para diferentes formas de expressar "Excluir" e "Tipo de dispositivo"
        padroes = {
            "Teclado": r'(Remover|Retirar|Apagar|Eliminar|Excluir)\s+Teclado',
            "Mouse": r'(Remover|Retirar|Apagar|Eliminar|Excluir)\s+Mouse',
            "Adaptador": r'(Remover|Retirar|Apagar|Eliminar|Excluir)\s+Adaptador',
            "Displayport": r'(Remover|Retirar|Apagar|Eliminar|Excluir)\s+Displayport',
            "Monitor": r'(Remover|Retirar|Apagar|Eliminar|Excluir)\s+Monitor'
        }

        if serial_dispositivo:
            # Carrega a planilha de estoque
            df = pd.read_excel(f"{CAMINHO_REDE}\\Estoque atual.xlsx")

            # Verifica se o serial está na planilha
            if serial_dispositivo in df['SerialDispositivo'].values:
                # Remove a linha correspondente ao serial
                df = df[df['SerialDispositivo'] != serial_dispositivo]
                # Salva as alterações na planilha
                df.to_excel(f"{CAMINHO_REDE}\\Estoque atual.xlsx", index=False)

                return f"Boa! Dispositivo com serial {serial_dispositivo} removido do estoque."
            else:
                return f"Poxa, o serial {serial_dispositivo} não foi encontrado no estoque."
        else:
            for tipo, padrao in padroes.items():
                match = re.search(padrao, mensagem, re.IGNORECASE)
                if match:
                    # Abrir campo para inserir o ID
                    id_dispositivo = st.text_input(f"Insira o ID do {tipo} a ser excluído:", key=f"ID_{tipo}")

                    if id_dispositivo:
                        # Carrega a planilha de estoque
                        df = pd.read_excel(f"{CAMINHO_REDE}\\Estoque atual.xlsx")
                        # Verifica se o ID está na planilha
                        if id_dispositivo in df['SerialDispositivo'].values:
                            # Remove a linha correspondente ao ID
                            df = df[df['SerialDispositivo'] != id_dispositivo]
                            # Salva as alterações na planilha
                            df.to_excel(f"{CAMINHO_REDE}\\Estoque atual.xlsx", index=False)

                            return f"Boa! {tipo} com ID {id_dispositivo} removido do estoque."
                        else:
                            return f"Poxa, o ID {id_dispositivo} do {tipo} não foi encontrado no estoque."

        return "Palavra-chave para exclusão não encontrada na mensagem."
    except Exception as e:
        return f"Erro ao excluir linha da planilha de estoque: {e}"

# Verifica o número de itens para um modelo específico
def verificarNumeroItensModelo(ModeloDispositivo):
    try:
        df = pd.read_excel(f"{CAMINHO_REDE}\\Estoque atual.xlsx")
        if ModeloDispositivo:
            num_itens = df[df['ModeloDispositivo'] == ModeloDispositivo].shape[0]
            return num_itens
        else:
            return None
    except Exception as e:
        return f"Erro ao verificar o número de itens para o modelo {ModeloDispositivo}: {e}"

# Função principal para processar a solicitação
async def main():
    mensagem = st.text_input("Adicione aqui o equipamento", key="text_input")
    if mensagem:
        tipos_dispositivos = ["Notebook", "Desktop", "Monitor", "Adaptador", "Displayport", "Teclado", "Mouse"]

        serial = verificSerial(mensagem)

        if re.search(r'(Excluir|Remover|Retirar|Apagar|Eliminar)', mensagem, re.IGNORECASE):
            resultado_exclusao = exclDados(mensagem, serial)
            st.write(resultado_exclusao)
        elif serial:
            # Consulta a planilha de estoque
            resultado, NomeDispositivo, ModeloDispositivo, SerialDispositivo, ProcessadorUsado, MemoriaTotal, ArmazenamentoInterno, ObservacaoDispositivo = consultDados(serial)
            if isinstance(resultado, pd.DataFrame):
                # Se os resultados forem retornados, mostra na interface
                st.write(resultado)
                # Pede para adicionar uma nova observação
                TipoDispositivo = st.selectbox("Tipo do Dispositivo:", tipos_dispositivos)
                ObservacaoDispositivo = st.text_input("Observação do Dispositivo:", key="ObservacaoDispositivo")
                if st.button("Enviar"):
                    # Adiciona a nova linha à planilha de estoque
                    resultado_adicao = adicDados(NomeDispositivo, ModeloDispositivo, SerialDispositivo, ProcessadorUsado, MemoriaTotal, ArmazenamentoInterno, ObservacaoDispositivo, TipoDispositivo)
                    st.write(resultado_adicao)
            else:
                st.write(resultado)
        else:
            for tipo in tipos_dispositivos:
                if tipo in mensagem:
                    Marca = st.text_input("Marca:", key="Marca")
                    ID = st.text_input("ID:", key="ID")
                    ObservacaoDispositivo = st.text_input("Observação do Dispositivo:", key="ObservacaoDispositivo")
                    TipoDispositivo = st.selectbox("Tipo do Dispositivo:", tipos_dispositivos)
                    if st.button("Adicionar ao estoque"):
                        resultado_adicao = adicDados(tipo, Marca, ID, None, None, None, ObservacaoDispositivo, TipoDispositivo)
                        st.write(resultado_adicao)
                    break
            else:
                st.write("Puts, nenhum equipamento foi encontrado com esse serial.")

# Executar a função principal
if __name__ == "__main__":
    asyncio.run(main())

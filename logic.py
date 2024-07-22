import re
import asyncio
import streamlit as st
import pandas as pd

# Verifica se a mensagem contém um padrão de número de série
def verificSerial(mensagem):
    padrao = r'\b[A-Za-z0-9]+\b'
    match = re.findall(padrao, mensagem)
    if match:
        return match[-1]  # Retorna o último elemento encontrado, que deve ser o número de série
    else:
        return None

# Consulta a planilha local em busca do número de série
def consultDados(Serial):
    try:
        df = pd.read_excel("C:\\Users\\lukin\\Downloads\\Scripts\\Goup-AutomateBot\\base\\Relatório Personalizado 14-05-2024.xlsx")
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
        
        df = pd.read_excel("C:\\Users\\lukin\\Downloads\\Scripts\\Goup-AutomateBot\\base\\Estoque atual.xlsx")
        df = pd.concat([df, nova_linha], ignore_index=True)
        df.to_excel("C:\\Users\\lukin\\Downloads\\Scripts\\Goup-AutomateBot\\base\\Estoque atual.xlsx", index=False)
        return "Boa bixo! Equipamento adicionado no estoque."
    except Exception as e:
        return f"Erro ao adicionar nova linha à planilha de estoque: {e}"

# Exclui uma linha da planilha de estoque com base no tipo de dispositivo e ID
def exclDados(mensagem):
    try:
        # Padroes para diferentes formas de expressar "Excluir" e "Tipo de dispositivo"
        padroes = {
            "Teclado": r'(?:Remover|Apagar|Eliminar|Excluir|Retirar)\s+Teclado',
            "Mouse": r'(?:Remover|Apagar|Eliminar|Excluir|Retirar)\s+Mouse',
            "Adaptador": r'(?:Remover|Apagar|Eliminar|Excluir|Retirar)\s+Adaptador',
            "Displayport": r'(?:Remover|Apagar|Eliminar|Excluir|Retirar)\s+Displayport',
            "Kit": r'(?:Remover|Apagar|Eliminar|Excluir|Retirar)\s+Kit'
        }

        for tipo, padrao in padroes.items():
            match = re.search(padrao, mensagem, re.IGNORECASE)
            if match:
                # Abrir campo para inserir o ID
                id_dispositivo = st.text_input(f"Insira o ID do {tipo} a ser excluído:", key=f"ID_{tipo}")
                if id_dispositivo:
                    # Carrega a planilha de estoque
                    df = pd.read_excel("C:\\Users\\lukin\\Downloads\\Scripts\\Goup-AutomateBot\\base\\Estoque atual.xlsx")

                    # Verifica se o ID está na planilha para o tipo de dispositivo especificado
                    if id_dispositivo in df[df['TipoDispositivo'] == tipo]['ID'].values:
                        # Remove a linha correspondente ao ID
                        df = df[(df['TipoDispositivo'] != tipo) | (df['ID'] != id_dispositivo)]
                        # Salva as alterações na planilha
                        df.to_excel("C:\\Users\\lukin\\Downloads\\Scripts\\Goup-AutomateBot\\base\\Estoque atual.xlsx", index=False)

                        return f"Boa! {tipo} com ID {id_dispositivo} removido do estoque."
                    else:
                        return f"Poxa, o ID {id_dispositivo} do {tipo} não foi encontrado no estoque."

        # Verifica se a mensagem contém um número de série
        serial = verificSerial(mensagem)
        if serial:
            # Carrega a planilha de estoque
            df = pd.read_excel("C:\\Users\\lukin\\Downloads\\Scripts\\Goup-AutomateBot\\base\\Estoque atual.xlsx")

            # Verifica se o número de série está na planilha
            if serial in df['SerialDispositivo'].values:
                # Remove a linha correspondente ao número de série
                df = df[df['SerialDispositivo'] != serial]
                # Salva as alterações na planilha
                df.to_excel("C:\\Users\\lukin\\Downloads\\Scripts\\Goup-AutomateBot\\base\\Estoque atual.xlsx", index=False)

                return f"Boa! O dispositivo com serial {serial} foi removido do estoque."
            else:
                return f"Poxa, o dispositivo com serial {serial} não foi encontrado no estoque."

        return "Palavra-chave para exclusão não encontrada na mensagem."
    except Exception as e:
        return f"Erro ao excluir linha da planilha de estoque: {e}"

# Verifica o número de itens para um modelo específico
def verificarNumeroItensModelo(ModeloDispositivo):
    try:
        df = pd.read_excel("C:\\Users\\lukin\\Downloads\\Scripts\\Goup-AutomateBot\\base\\Estoque atual.xlsx")
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
        
        if any(word in mensagem.lower() for word in ["excluir", "remover", "retirar", "apagar", "eliminar"]):
            resultado_exclusao = exclDados(mensagem)
            st.write(resultado_exclusao)
        else:
            Serial = verificSerial(mensagem)
            if Serial:
                # Consulta a planilha de estoque
                resultado, NomeDispositivo, ModeloDispositivo, SerialDispositivo, ProcessadorUsado, MemoriaTotal, ArmazenamentoInterno, ObservacaoDispositivo = consultDados(Serial)
                if isinstance(resultado, pd.DataFrame):
                    # Se os resultados forem retornados, mostra na interface
                    st.write(resultado)
                    # Pede para adicionar uma nova observação
                    TipoDispositivo = st.selectbox("Tipo do Dispositivo:", tipos_dispositivos)  # Campo de seleção para o tipo do dispositivo
                    ObservacaoDispositivo = st.text_input("Observação do Dispositivo:", key="ObservacaoDispositivo")
                    if st.button("Enviar"):
                        # Adiciona a nova linha à planilha de estoque
                        resultado_adicao = adicDados(NomeDispositivo, ModeloDispositivo, SerialDispositivo, ProcessadorUsado, MemoriaTotal, ArmazenamentoInterno, ObservacaoDispositivo, TipoDispositivo)
                        st.write(resultado_adicao)
                else:
                    st.write(resultado)
            else:
                st.write("Puts, nenhum equipamento foi encontrado com esse serial.")

# Executar a função principal
if __name__ == "__main__":
    asyncio.run(main())

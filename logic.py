import re
import asyncio
import streamlit as st
import pandas as pd

# Verifica se a mensagem contém um padrão de número de série
def verificSerial(mensagem):
    padrao = r'\b[A-Za-z0-9]+\b'
    match = re.search(padrao, mensagem)
    if match:
        return match.group()
    else:
        return None

# Consulta a planilha local em busca do número de série
def consultDados(Serial):
    try:
        df = pd.read_excel(r"\\10.111.200.59\Controle\Estoque\Relatório Personalizado 14-05-2024.xlsx")
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
        
        df = pd.read_excel(r"\\10.111.200.59\Controle\Estoque\Estoque atual.xlsx")
        df = pd.concat([df, nova_linha], ignore_index=True)
        df.to_excel(r"\\10.111.200.59\Controle\Estoque\Estoque atual.xlsx", index=False)
        return "Boa bixo! Equipamento adicionado no estoque."
    except Exception as e:
        return f"Erro ao adicionar nova linha à planilha de estoque: {e}"

# Exclui uma linha da planilha de estoque com base no número de série
def exclDados(mensagem):
    try:
        # Verifica se a mensagem contém a palavra "Excluir" e extrai o número de série
        padrao = r'Excluir\s+(?P<Serial>\S+)'
        match = re.search(padrao, mensagem)
        if match:
            Serial = match.group('Serial')
            
            # Carrega a planilha de estoque
            df = pd.read_excel(r"\\10.111.200.59\Controle\Estoque\Estoque atual.xlsx")
            
            # Verifica se o número de série está na planilha
            if Serial in df['SerialDispositivo'].values:
                # Obtém o modelo do item excluído
                modelo_excluido = df[df['SerialDispositivo'] == Serial]['ModeloDispositivo'].iloc[0]
                
                # Remove a linha correspondente ao número de série
                df = df[df['SerialDispositivo'] != Serial]
                # Salva as alterações na planilha
                df.to_excel(r"\\10.111.200.59\Controle\Estoque\Estoque atual.xlsx", index=False)
                
                # Verifica se o modelo excluído tem 5 ou menos itens no estoque
                num_itens_modelo = verificarNumeroItensModelo(modelo_excluido)
                if num_itens_modelo is not None:
                    if num_itens_modelo == 0:
                        st.error(f"Atenção: O modelo {modelo_excluido} agora está esgotado!")
                    elif num_itens_modelo <= 5:
                        st.warning(f"Atenção: O modelo {modelo_excluido} agora tem apenas {num_itens_modelo} itens no estoque.")
                
                return "Boa! Equipamento removido do estoque."
            else:
                return f"Poxa, o Serial - '{Serial}' não foi encontrado no estoque."
        else:
            return "Palavra-chave 'Excluir' não encontrada na mensagem."
    except Exception as e:
        return f"Erro ao excluir linha da planilha de estoque: {e}"

# Verifica o número de itens para um modelo específico
def verificarNumeroItensModelo(ModeloDispositivo):
    try:
        df = pd.read_excel(r"\\10.111.200.59\Controle\Estoque\Estoque atual.xlsx")
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
        
        if "Monitor" in mensagem:
            Marca = st.text_input("Marca:", key="Marca")
            ModeloDispositivo = st.text_input("Modelo do Monitor:", key="ModeloDispositivo")
            Serial = st.text_input("Serial:", key="Serial")
            ObservacaoDispositivo = st.text_input("Observação do Dispositivo:", key="ObservacaoDispositivo")
            TipoDispositivo = st.selectbox("Tipo do Dispositivo:", tipos_dispositivos)  # Campo de seleção para o tipo do dispositivo
            if st.button("Adicionar ao estoque"):
                resultado_adicao = adicDados("Monitor", ModeloDispositivo, Serial, None, None, None, ObservacaoDispositivo, TipoDispositivo)
                st.write(resultado_adicao)
        elif "Teclado" in mensagem:
            Marca = st.text_input("Marca:", key="Marca")
            ID = st.text_input("ID:", key="ID")
            ObservacaoDispositivo = st.text_input("Observação do Dispositivo:", key="ObservacaoDispositivo")
            TipoDispositivo = st.selectbox("Tipo do Dispositivo:", tipos_dispositivos)  # Campo de seleção para o tipo do dispositivo
            if st.button("Adicionar ao estoque"):
                resultado_adicao = adicDados("Teclado", Marca, ID, None, None, None, ObservacaoDispositivo, TipoDispositivo)
                st.write(resultado_adicao)
        elif "Mouse" in mensagem:
            Marca = st.text_input("Marca:", key="Marca")
            ID = st.text_input("ID:", key="ID")
            ObservacaoDispositivo = st.text_input("Observação do Dispositivo:", key="ObservacaoDispositivo")
            TipoDispositivo = st.selectbox("Tipo do Dispositivo:", tipos_dispositivos)  # Campo de seleção para o tipo do dispositivo
            if st.button("Adicionar ao estoque"):
                resultado_adicao = adicDados("Mouse", Marca, ID, None, None, None, ObservacaoDispositivo, TipoDispositivo)
                st.write(resultado_adicao)
        elif "Adaptador" in mensagem:
            ID = st.text_input("ID:", key="ID")
            ObservacaoDispositivo = st.text_input("Observação do Dispositivo:", key="ObservacaoDispositivo")
            TipoDispositivo = st.selectbox("Tipo do Dispositivo:", tipos_dispositivos)  # Campo de seleção para o tipo do dispositivo
            if st.button("Adicionar ao estoque"):
                resultado_adicao = adicDados("Adaptador", None, ID, None, None, None, ObservacaoDispositivo, TipoDispositivo)
                st.write(resultado_adicao)
        elif "Displayport" in mensagem:
            ID = st.text_input("ID:", key="ID")
            ObservacaoDispositivo = st.text_input("Observação do Dispositivo:", key="ObservacaoDispositivo")
            TipoDispositivo = st.selectbox("Tipo do Dispositivo:", tipos_dispositivos)  # Campo de seleção para o tipo do dispositivo
            if st.button("Adicionar ao estoque"):
                resultado_adicao = adicDados("Displayport", None, ID, None, None, None, ObservacaoDispositivo, TipoDispositivo)
                st.write(resultado_adicao)
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
                    if "Excluir" not in mensagem:  # Ignora a mensagem se a função de exclusão for acionada
                        st.write(resultado)
                
                # Verifica se a mensagem contém o comando de exclusão
                if "Excluir" in mensagem:
                    resultado_exclusao = exclDados(mensagem)
                    st.write(resultado_exclusao)
            else:
                st.write("Puts, nenhum equipamento foi encontrado com esse serial.")

# Executar a função principal
if __name__ == "__main__":
    asyncio.run(main())
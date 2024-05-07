import tkinter as tk
from tkinter import filedialog, Entry, Button, Label, StringVar, CENTER
import subprocess
import NFSe
import pandas as pd
import BancoDeDados
import os

def processar_arquivos_pasta():
    # Abre o diálogo para selecionar uma pasta
    pasta_selecionada = filedialog.askdirectory()
    if pasta_selecionada:
        # Lista todos os arquivos na pasta selecionada
        arquivos_na_pasta = os.listdir(pasta_selecionada)
        for arquivo in arquivos_na_pasta:
            # Verifica se o arquivo é um arquivo CSV
            if arquivo.endswith('.csv'):
                caminho_arquivo = os.path.join(pasta_selecionada, arquivo)
                # Aqui você pode fazer o que desejar com o caminho do arquivo,
                # como armazená-lo para uso futuro ou realizar outras operações
        return pasta_selecionada
def processar_arquivo_NFSe():
    # Verifica se a pasta selecionada é válida
    pasta_selecionada = processar_arquivos_pasta()
    if pasta_selecionada:
        # Lista todos os arquivos na pasta selecionada
        arquivos_na_pasta = os.listdir(pasta_selecionada)
        for arquivo in arquivos_na_pasta:
            # Verifica se o arquivo é um arquivo CSV
            if arquivo.endswith('.csv'):
                caminho_arquivo = os.path.join(pasta_selecionada, arquivo)
                # Usar apenas o nome do arquivo, não o caminho completo
                nome_arquivo = os.path.basename(caminho_arquivo)
                # Remover a extensão .csv ou .txt do nome do arquivo
                nome_formatado = nome_arquivo.replace(".csv", ".txt")

                # Processa o arquivo
                resultado_processamento = NFSe.processar_arquivoNF(caminho_arquivo)
                if BancoDeDados.campoTextoVerifica1 in resultado_processamento and BancoDeDados.campoTextoVerifica2 in resultado_processamento:
                    # Exibe a mensagem no rótulo
                    label_resultado.config(text=resultado_processamento)
                # Verifica se o resultado do processamento é válido
                elif resultado_processamento:
                    # Se não houver linhas relevantes, cria um novo arquivo com todas as linhas
                    caminho_pasta = os.path.dirname(caminho_arquivo)
                    caminho_arquivo_txt = os.path.join(caminho_pasta, nome_formatado)
                    with open(caminho_arquivo_txt, 'w', encoding='utf-8') as arquivo_saida:
                        arquivo_saida.write(resultado_processamento)

                    # Abre o arquivo no programa associado ao tipo de arquivo
                    # subprocess.Popen(["notepad.exe", caminho_arquivo_txt])

                    # Exibe a mensagem no rótulo
                    label_resultado.config(text=BancoDeDados.campoArquivoGerado)
                else:
                    # Exibe a mensagem de erro no rótulo
                    label_resultado.config(text="Erro ao processar arquivo: " + nome_arquivo)
    else:
        # Exibe a mensagem de erro no rótulo se a pasta não for selecionada
        label_resultado.config(text="Erro: Pasta não selecionada.")
def realizar_operacaoNFTS():
    # Verifica se a pasta selecionada é válida
    pasta_selecionada = processar_arquivos_pasta()
    if pasta_selecionada:
        # Lista todos os arquivos na pasta selecionada
        arquivos_na_pasta = os.listdir(pasta_selecionada)
        for arquivo in arquivos_na_pasta:
            # Verifica se o arquivo é um arquivo CSV
            if arquivo.endswith('.csv'):
                caminho_arquivo = os.path.join(pasta_selecionada, arquivo)
                # Usar apenas o nome do arquivo, não o caminho completo
                nome_arquivo = os.path.basename(caminho_arquivo)
                # Remover a extensão .csv ou .txt do nome do arquivo
                nome_formatado = nome_arquivo.replace(".csv", ".txt")

                # Processa o arquivo
                resultado_processamento = NFSe.processar_arquivoNFTS(caminho_arquivo)
                if resultado_processamento is None:
                    if BancoDeDados.campoTextoVerifica1 in resultado_processamento and BancoDeDados.campoTextoVerifica2 in resultado_processamento:
                        label_resultado.config(text=resultado_processamento)
                # Verifica se o resultado do processamento é válido
                elif resultado_processamento:
                    # Se não houver linhas relevantes, cria um novo arquivo com todas as linhas
                    caminho_pasta = os.path.dirname(caminho_arquivo)
                    caminho_arquivo_txt = os.path.join(caminho_pasta, nome_formatado)
                    with open(caminho_arquivo_txt, 'w', encoding='utf-8') as arquivo_saida:
                        arquivo_saida.write(resultado_processamento)

                    # Abre o arquivo no programa associado ao tipo de arquivo
                    # subprocess.Popen(["notepad.exe", caminho_arquivo_txt])

                    # Exibe a mensagem no rótulo
                    label_resultado.config(text=BancoDeDados.campoArquivoGerado)
                else:
                    # Exibe a mensagem de erro no rótulo
                    label_resultado.config(text="Erro ao processar arquivo: " + nome_arquivo)
    else:
        # Exibe a mensagem de erro no rótulo se a pasta não for selecionada
        label_resultado.config(text="Erro: Pasta não selecionada.")
def adicionar_acumulador():
    codigo = str(entry_codigo.get())
    natureza = str(entry_natureza.get())
    if codigo and natureza:
        try:
            # Força a coluna 'Cod' a ser tratada como string ao ler o arquivo Excel
            df = pd.read_excel(BancoDeDados.planilhaAcumuladores, dtype={BancoDeDados.campoColunaProcurada: str, BancoDeDados.campoColunaResultado: str})
        except FileNotFoundError:
            df = pd.DataFrame(columns=[BancoDeDados.campoColunaProcurada, BancoDeDados.campoColunaResultado])
        # Verifica se o código já existe no DataFrame
        if codigo in df[BancoDeDados.campoColunaProcurada].values:
            # Atualiza a natureza se o código já existe
            df.loc[df[BancoDeDados.campoColunaProcurada] == codigo, BancoDeDados.campoColunaResultado] = natureza
        else:
            # Adiciona uma nova entrada se o código não existe
            nova_entrada = pd.DataFrame({BancoDeDados.campoColunaProcurada: [codigo], BancoDeDados.campoColunaResultado: [natureza]})
            df = pd.concat([df, nova_entrada], ignore_index=True)
        # Força a coluna 'Cod' a ser tratada como string ao salvar o DataFrame no Excel
        df[BancoDeDados.campoColunaProcurada] = df[BancoDeDados.campoColunaProcurada].astype(str)
        df.to_excel(BancoDeDados.planilhaAcumuladores, index=False)
        # Atualiza a exibição na interface formatando o valor como string
        label_resultado.config(text=f"{BancoDeDados.campoDadosCadastrados} {natureza}")
    else:
        label_resultado.config(text=BancoDeDados.campoNaoPreencheu)
def mostrar_ocultar_cadastro():
    global altura_janela
    nova_altura = 450
    janela.geometry(f"{largura_janela}x{nova_altura}")
    # Função chamada quando o botão "Cadastrar Natureza de Rendimento" é clicado
    if frame_cadastro.winfo_ismapped():
        frame_cadastro.pack_forget()
    else:
        frame_cadastro.pack()
# Criação da janela principal
janela = tk.Tk()
janela.title(BancoDeDados.nomeDoProgama)
# Adicionando um logotipo
logo_path = BancoDeDados.caminhoLogo  # Substitua pelo caminho do seu logotipo
logo = tk.PhotoImage(file=logo_path)
# Ajustando o tamanho do logotipo para 100x100 pixels usando subsample
logo = logo.subsample(8)  # Pode precisar ajustar o valor de subsample
logo_label = Label(janela, image=logo)
logo_label.pack(pady=0)
var_caminho_arquivo = StringVar()
var_caminho_salvar = StringVar()
# Rótulo para exibir o caminho do arquivo selecionado
#rotulo_caminho_arquivo = tk.Label(janela, textvariable=var_caminho_arquivo)
#rotulo_caminho_arquivo.pack(pady=9)
largura_janela = BancoDeDados.tamanhoPaginaX
altura_janela = BancoDeDados.tamanhoPaginaY
# Configurando a geometria da janela
janela.geometry(f"{largura_janela}x{altura_janela}")
# Variáveis para armazenar os caminhos do arquivo e de salvamento
var_caminho_arquivo = tk.StringVar()
# Botao Selecionar Arquivo
#botao_selecionar_arquivo = Button(janela, text=BancoDeDados.nomeBotaoSelecionar, command=processar_arquivos_pasta)
#botao_selecionar_arquivo.pack(pady=(10,5))

# Botao Gerar NF-s
botao_ok = Button(janela, text=BancoDeDados.nomeBotaoGerar, command=processar_arquivo_NFSe)
botao_ok.pack(side="top", pady=5)
# Botao Gerar NFTS
botao_NFTS = Button(janela, text=BancoDeDados.nomeBotaoGerarNFTS, command=realizar_operacaoNFTS)
botao_NFTS.pack(side="top", pady=5)
# Botão para mostrar ou ocultar o frame de cadastro
botao_cadastrar = Button(janela, text=BancoDeDados.nomeBotaoCadastrar, command=mostrar_ocultar_cadastro)
botao_cadastrar.pack(side="top", anchor="center", pady=5)

# Botao Criar campos de cadastro caso seja clicado!
frame_cadastro = tk.Frame(janela)

criador = Label(janela, text=BancoDeDados.Criador, font=("Arial",7))
criador.pack(side="top", pady=5)

# Criar campos de entrada no frame de cadastro
label_codigo = Label(frame_cadastro, text=BancoDeDados.nomeCadastroCod)
label_codigo.grid(row=0, column=0, padx=10, pady=5)
entry_codigo = Entry(frame_cadastro)
entry_codigo.grid(row=0, column=1, padx=10, pady=5)
label_natureza = Label(frame_cadastro, text=BancoDeDados.nomeCadastroNat)
label_natureza.grid(row=1, column=0, padx=10, pady=5)
entry_natureza = Entry(frame_cadastro)
entry_natureza.grid(row=1, column=1, padx=10, pady=5)
# Botão para adicionar acumulador
button_adicionar = Button(frame_cadastro, text=BancoDeDados.nomeCadastrar, command=adicionar_acumulador)
button_adicionar.grid(row=2, column=0, columnspan=2, pady=5)
# Rótulo para exibir o resultado
label_resultado = Label(janela, text="")
label_resultado.pack(pady=5)

# Inicia o loop principal da interface gráfica
janela.mainloop()

import os
import pandas as pd
import tkinter as tk
from tkinter import messagebox
import numpy as np
import BancoDeDados


def fazerProcv(codServ):
    df = pd.read_excel(BancoDeDados.planilhaAcumuladores)
    if not df.empty and codServ in df[BancoDeDados.campoColunaProcurada].values:
        resultado = str(df.loc[df[BancoDeDados.campoColunaProcurada] == codServ, BancoDeDados.campoColunaAcumulador].values[0])
        return resultado
    else:
        # Lida com a situação em que o código de serviço não é encontrado
        return ""

def Procv1020(codServ):
    df = pd.read_excel(BancoDeDados.planilhaAcumuladores)
    if not df.empty and codServ in df[BancoDeDados.campoColunaProcurada].values:
        resultado = str(df.loc[df[BancoDeDados.campoColunaProcurada] == codServ, BancoDeDados.campoColunaGera].values[0])
        return resultado
    else:
        # Lida com a situação em que o código de serviço não é encontrado
        return ""

def procvIR(codServ):
    df = pd.read_excel(BancoDeDados.planilhaAcumuladores, dtype={BancoDeDados.campoColunaIR: str})
    if not df.empty and codServ in df[BancoDeDados.campoColunaProcurada].values:
        resultado = str(df.loc[df[BancoDeDados.campoColunaProcurada] == codServ, BancoDeDados.campoColunaIR].values[0])
        return resultado
    else:
        # Lida com a situação em que o código de serviço não é encontrado
        return ""
def procvCRF(codServ):
    df = pd.read_excel(BancoDeDados.planilhaAcumuladores)
    if not df.empty and codServ in df[BancoDeDados.campoColunaProcurada].values:
        resultado = str(df.loc[df[BancoDeDados.campoColunaProcurada] == codServ, BancoDeDados.campoColunaCRF].values[0])
        return resultado
    else:
        # Lida com a situação em que o código de serviço não é encontrado
        return ""
def procvNatRendimento(codServ):

    df = pd.read_excel(BancoDeDados.planilhaAcumuladores, dtype={BancoDeDados.campoColunaResultado: str})
    if not df.empty:
        if 'Cod' in df.columns:
            # Verifica se codServ está presente na coluna 'Cod'
            if df[BancoDeDados.campoColunaProcurada].eq(codServ).any():
                resultado = str(df.loc[df[BancoDeDados.campoColunaProcurada] == codServ, BancoDeDados.campoColunaResultado].values[0])
                resultado = pd.to_numeric(resultado, errors='coerce')
                if pd.isna(resultado):
                    mensagem = f"{BancoDeDados.campoTextoVerifica1} {codServ} \n{BancoDeDados.campoTextoVerifica2}."
                    messagebox.showwarning("Aviso", mensagem)
                    return ''
                else:
                    return resultado
            else:
                # Se codServ não está presente na coluna 'Cod'
                mensagem = f"Codigo de Serviço {codServ} não cadastrado!"
                messagebox.showwarning("Aviso", mensagem)
                return ''

        else:
            mensagem = "A coluna 'Cod' não está presente no DataFrame."
            messagebox.showerror("Erro", mensagem)
            return ''
    else:
        mensagem = "O DataFrame está vazio. Verifique o arquivo de dados."
        messagebox.showerror("Erro", mensagem)
        return ''


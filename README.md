# 🌀 Loading Spinner Moderno e Reutilizável no Visual Basic 6 (VB6)

Este projeto ensina como criar um **Loading Spinner moderno e reutilizável** no Visual Basic 6 (VB6), ideal para indicar carregamentos em suas aplicações.

## ✨ Recursos
- Design moderno e profissional
- Fácil integração e reutilização em vários projetos VB6
- Animação suave para melhor experiência do usuário

## 📌 Como Utilizar

### 1️⃣ Importar o Formulário de Spinner  
Baixe o arquivo **`frmSpinner.frm`** e adicione-o ao seu projeto VB6.

### 2️⃣ Criar a Função para Exibir o Spinner  
No seu formulário principal, adicione a seguinte função:

```vb
Dim Spinner As New frmSpinner

Public Sub ShowSpinner()
    Spinner.Show vbModeless
    DoEvents
End Sub

Public Sub HideSpinner()
    Unload Spinner
End Sub

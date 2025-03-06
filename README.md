# 🌀 Loading Spinner Moderno e Reutilizável no Visual Basic 6 (VB6)

Este projeto ensina como criar um **Loading Spinner moderno e reutilizável** no Visual Basic 6 (VB6), ideal para indicar carregamentos em suas aplicações.

## ✨ Recursos
- Design moderno e profissional
- Fácil integração e reutilização em vários projetos VB6
- Animação suave para melhor experiência do usuário

## 📌 Como Utilizar

### 1️⃣ Importar o Formulário de Spinner  
Baixe os arquivos **`frm_loading_spinner.frm`** / **`ModuloLoadingSpinner.bas`** adicione-o ao seu projeto VB6.

### 2️⃣ Chamar a Sub para Exibir o Spinner
Use "IsLoading True" no inicio do processamento
Use "DoEvents" durante o processamento
Use "IsLoading False" no final do processamento
No seu formulário principal ou em algum formulário que possua um processamento demorado (exemplo ao clicar em um botão), chame a Sub da seguinte maneira:

```vb
Private Sub Command1_Click()

   IsLoading True

   Dim I As Long
   For I = 1 To 9999999
    DoEvents
   Next I

   IsLoading False

End Sub

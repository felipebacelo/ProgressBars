![GitHub repo size](https://img.shields.io/github/repo-size/felipebacelo/ProgressBars?style=for-the-badge)
![GitHub language count](https://img.shields.io/github/languages/count/felipebacelo/ProgressBars?style=for-the-badge)
![GitHub forks](https://img.shields.io/github/forks/felipebacelo/ProgressBars?style=for-the-badge)
![Bitbucket open pull requests](https://img.shields.io/bitbucket/pr-raw/felipebacelo/ProgressBars?style=for-the-badge)
![Bitbucket open issues](https://img.shields.io/bitbucket/issues/felipebacelo/ProgressBars?style=for-the-badge)

# ProgressBars
ProgressBars - VBA Excel

Simples exemplo de como podemos criar uma barra de progresso através do VBA.

### Desenvolvimento

Desenvolvido em Microsoft VBA Excel.
***
### Requisitos

* Habilitar Macros
* Habilitar Guia de Desenvolvedor

### Referências às Bibliotecas

* Visual Basic For Applications
* Microsoft Excel 16.0 Object Library
* OLE Automation
* Microsoft Office 16.0 Object Library
* Microsoft Forms 2.0 Object Library

### Compatibilidade

Este exemplo foi desenvolvido no Excel 2019 (64 bits) e testado no Excel 2016 (64 bits). Sua compatibilidade é garantida para a versão 2016 e superiores. Sua utilização em versões anteriores pode ocasionar em não funcionamento do mesmo.

### Usabilidade

Para utilizar este exemplo o usuário deverá:

* Realizar o download do arquivo ZIP: __ProgressBars__.
* Abrir o arquivo __ProgressBars.xlsm__, ou importar através do VBA os arquivos __Módulo1.bas__ e __UserForm1.frm__.
***
### Demo

![GIF](https://github.com/felipebacelo/ProgressBars/blob/main/Demo.gif)

***
### Exemplo de Macro Utilizada

```vba
Option Explicit

Private Sub UserForm_Activate()

ProgressBar.Width = 0

Do While ProgressBar.Width < 396
    
    Sleep (10)

    ProgressBar.Width = ProgressBar.Width + 2
    
    DoEvents
    
Loop

MsgBox "Seja Bem Vindo ao ProgressBar!!!", vbInformation, "ProgressBar"

Me.Hide

End Sub
```
***
### Licenças

_MIT License_
_Copyright   ©   2020 Felipe Bacelo Rodrigues_


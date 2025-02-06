# Langkah 1: Instalasi Perangkat Lunak
Instalasi perangkat lunak yang diperlukan, seperti:

- Access 3.2
- Node.js
- TON Dev Suite
- Library ODBC untuk Access

# Langkah 2: Buat Koneksi ODBC
Buat koneksi ODBC untuk menghubungkan Access dengan TON Blockchain:

- Buka Control Panel > Administrative Tools > Data Sources (ODBC)
- Klik "Add" dan pilih "TON Blockchain ODBC Driver"
- Isi informasi koneksi, seperti alamat node TON, port, dan kredensial

# Langkah 3: Buat Code VBA di Access
Buat code VBA di Access untuk menghubungkan dengan TON Blockchain:

- Buka Access dan buka modul VBA
- Tambahkan referensi ke library ODBC
- Tulis code VBA untuk menghubungkan dengan TON Blockchain menggunakan koneksi ODBC

Contoh code VBA:

```
Sub ConnectToTON(https://t.me/UXUYbot/app?startapp=B_0x25A71bEa3659d81CA45CcAAb87b146e40d683Ec8bnbchain0x55d398326f99059fF775485246999027B3197955)
    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    
    ' Buat koneksi ODBC
    Set cn = New ADODB.Connection
    cn.Open "DSN=TON Blockchain ODBC Driver;UID=myuser;PWD=mypassword"
    
    ' Buat recordset untuk menyimpan data
    Set rs = New ADODB.Recordset
    rs.Open "SELECT * FROM mytable", cn
    
    ' Lakukan operasi pada data
    ' ...
    
    ' Tutup recordset dan koneksi
    rs.Close
    cn.Close
    
    Set rs = Nothing
    Set cn = Nothing
End Sub
```

# Langkah 4: Integrasi dengan TON Blockchain
Integrasi code VBA dengan TON Blockchain menggunakan library TON Dev Suite:

- Tambahkan referensi ke library TON Dev Suite
- Tulis code VBA untuk mengirimkan data ke TON Blockchain menggunakan library TON Dev Suite

Contoh code VBA:

```
Sub SendDataToTON(https://t.me/UXUYbot/app?startapp=B_0x25A71bEa3659d81CA45CcAAb87b146e40d683Ec8bnbchain0x55d398326f99059fF775485246999027B3197955)
    Dim ton As New TONDevSuite.TON
    Dim data As String
    
    ' Siapkan data untuk dikirim
    data = "Hello, TON Blockchain!"
    
    ' Kirim data ke TON Blockchain
    ton.SendMessage data
    
    ' Tutup koneksi
    ton.Close
End Sub
```
# GitHub Copilot extension: Overview

[GitHub Copilot](https://github.com/features/copilot) is an AI-powered pair programmer extension for [!INCLUDE [azure-data-studio-short](../includes/azure-data-studio-short.md)] that provides you with context-aware code completions, suggestions, and even entire code snippets. This powerful tool helps developers write code more efficiently, reduce the time spent on repetitive tasks, and minimize errors.

## What is GitHub Copilot?

GitHub Copilot for [!INCLUDE [azure-data-studio-short](../includes/azure-data-studio-short.md)] can be used in any editor window. To use GitHub Copilot, you must have an active internet connection. You can use GitHub Copilot in the following ways:

1. When you type code in the editor, GitHub Copilot provides suggestions in real-time.
1. When you type a natural language comment, GitHub Copilot provides suggestions for code that corresponds to the comment.

To accept a suggestion, press `Tab`. To reject a suggestion, press `Esc`.

At any time, pressing `Ctrl`+`Enter` opens the GitHub Copilot Completions Panel, which provides suggestions for code based on the context of the editor.

GitHub Copilot chat isn't currently available for [!INCLUDE [azure-data-studio-short](../includes/azure-data-studio-short.md)].

## Install the GitHub Copilot extension

To get started, all you need is [!INCLUDE [azure-data-studio-short](../includes/azure-data-studio-short.md)] [version 1.44](../release-notes-azure-data-studio.md#may-2023) or later, and a GitHub Copilot [subscription](https://docs.github.com/en/enterprise-cloud@latest/billing/managing-billing-for-github-copilot/about-billing-for-github-copilot).

> [!TIP]  
> GitHub Copilot is free for verified students and for maintainers of popular open source projects on GitHub.

1. Select the Extensions Icon to view the available extensions.

   :::image type="content" source="media/add-extensions/extension-manager-icon.png" alt-text="Screenshot showing the Extension manager icon.":::

1. Search for the **GitHub Copilot** extension and select it to view its details. Select **Install** to add the extension.

## How GitHub Copilot works

GitHub Copilot works by utilizing advanced machine learning models trained on a vast dataset of publicly available code from GitHub repositories. As you type code, the AI analyzes the context and provides relevant suggestions in real-time. You can receive suggestions also by writing a natural language comment describing what you want the code to do.

The GitHub Copilot extension in [!INCLUDE [azure-data-studio-short](../includes/azure-data-studio-short.md)] uses context from the editor to provide suggestions. For example, if you're writing a query that joins two tables, GitHub Copilot suggests the join condition from columns in the open editor, other files in the workspace, and common syntax patterns.

## Privacy

Your code is yours. We follow responsible practices in accordance with our [Privacy Statement](https://docs.github.com/site-policy/privacy-policies/github-privacy-statement) to ensure that your code snippets aren't used as suggested code for other users of GitHub Copilot.

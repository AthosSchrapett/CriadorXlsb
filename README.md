# 📂 CriadorXLSB

[![NuGet](https://img.shields.io/nuget/v/CriadorXLSB)](https://www.nuget.org/packages/CriadorXLSB/)
[![License](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE)

**CriadorXLSB** é uma biblioteca C# que permite converter arquivos **XLSX** para **XLSB** utilizando Python.
Ideal para automação de planilhas e melhoria de performance ao lidar com grandes volumes de dados no Excel.

---

## 📌 **Instalação**
Você pode instalar a biblioteca diretamente do **NuGet**:

```sh
dotnet add package CriadorXLSB --version 1.0.6
```

Ou via **Gerenciador de Pacotes NuGet** no Visual Studio.

---

## 🚀 **Como Usar**
### **1️⃣ Criar um Objeto de Caminhos**
Antes de iniciar, crie um objeto **Caminhos** para armazenar os caminhos do script e arquivos:

```csharp
using CriadorXLSB;
string input = @"seu caminho de entrada/dados.xlsx"
string output = @"seu caminho de saida/dados.xlsb"
var caminhos = new Caminhos(input, output);
```

---

### **2️⃣ Chamar o Conversor**
Agora, basta chamar o método de conversão:

```csharp
var gerador = new GeradorXlsb(caminhos);
gerador.GerarXlsb();
```

Isso executará o script Python e gerará o arquivo XLSB automaticamente.

---

## 📂 **Estrutura do Projeto**
```
📦 CriadorXLSB
 ┣ 📜 Caminhos.cs       # Define os caminhos do script e arquivos
 ┣ 📜 GeradorXlsb.cs    # Classe principal para conversão
 ┣ 📜 PythonHelper.cs   # Executa o script Python no C#
 ┣ 📜 converterXlsxEmXlsb.py # Script Python para conversão
```

---

## 🛠 **Requisitos**
- **Python 3.13+** instalado no sistema.
- Bibliotecas Python necessárias:
  ```sh
  pip install pandas pywin32
  ```
- **Microsoft Excel instalado** no sistema (requisito para o COM API do Windows).

---

## 📜 **Licença**
Este projeto é licenciado sob a licença MIT - veja o arquivo [LICENSE](LICENSE) para mais detalhes.

---

## 🌟 **Contribuições**
Sinta-se à vontade para contribuir! Se encontrar problemas ou quiser adicionar melhorias, abra uma **issue** ou envie um **pull request**.

📬 **Contato:** [AthosSchrapett](https://github.com/AthosSchrapett)

---

## 🏆 **Dúvidas ou Problemas?**
Caso tenha algum problema ao utilizar o CriadorXLSB, verifique:
1. Se o **Python está instalado** e acessível no sistema.
2. Se as bibliotecas necessárias foram instaladas (`pip install pandas pywin32`).
3. Se o **arquivo Python está presente na pasta `bin/Debug/net9.0/`** após a instalação do pacote.
4. Se o Excel está corretamente instalado no sistema.

Se precisar de ajuda, abra uma **issue no GitHub**! 🚀


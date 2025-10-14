# Central PPT Maker

A desktop application to create PowerPoint presentations for church services.

---

## Description

**Central PPT Slide Maker** is a user-friendly tool designed to streamline the creation of PowerPoint presentations for church services. It automates the process of formatting song lyrics, service sections, and slide layouts, saving time and reducing manual effort. 

This application turned a 2 hour task into a 15 minute task.

---

## Motivation

Creating PowerPoint presentations for church services can be repetitive and time-consuming. This application simplifies the process by providing a straightforward interface for entering the order of service and automatically generating a polished presentation.

---

## 🌟 Features

- **📝 File Name Input** - Set a custom file name for your presentation
- **⛪ Service Type Selection** - Choose between **Sunday** or **Midweek** services
- **📋 Order of Service Input** - Paste song lyrics and service details into the text box
- **🎨 Text Formatting Tools** - Buttons for undo, redo, copy, paste, and formatting text as song titles
- **🔄 Slide Creation** - Use keywords to create different types of slides (welcome, communion, sermon, etc.)
- **🧹 File Sanitization** - Automatically removes invalid characters from file names
- **❓ Help Dialog** - Detailed instructions on how to use the application

---

## 🛠 Tech Stack

![Python](https://img.shields.io/badge/Python-3776AB?style=for-the-badge&logo=python&logoColor=white)
![Tkinter](https://img.shields.io/badge/Tkinter-3776AB?style=for-the-badge&logo=python&logoColor=white)
![PowerPoint](https://img.shields.io/badge/Microsoft_PowerPoint-B7472A?style=for-the-badge&logo=microsoft-powerpoint&logoColor=white)
![Poetry](https://img.shields.io/badge/Poetry-60A5FA?style=for-the-badge&logo=poetry&logoColor=white)

---

## 🚀 Quick Start

### Prerequisites
- Python 3.x
- Poetry

### Installation & Usage

1. **Clone and setup**:
   ```bash
   git clone <repository-url>
   cd <repository-directory>
   poetry install
   ```

2. **Run the application**:
   ```bash
   poetry run python main.py
   ```

3. **Create your presentation**:
   - Enter a file name
   - Select service type (Sunday/Midweek)
   - Paste order of service with keywords
   - Generate your PowerPoint!

---

### 🎯 Keywords for Slide Creation

| Keyword | Slide Type |
|---------|------------|
| `WELCOME/PRAYER` | Welcome/Prayer slide |
| `COMMUNION` | Communion slide |
| `SERMON` | Sermon slide |
| `CLOSE` | Closing slide |
| `CONTRIBUTION` | Contribution slide |
| `TITLE` | Song title slide |

---

### 📋 Example Input

```plaintext
WELCOME/PRAYER (John Doe)

TITLE (Amazing Grace)

Amazing grace! how sweet the sound,
That saved a wretch like me!
I once was lost but now am found,
Was blind, but now I see.

'Tis grace that taught my heart to fear,
And grace my fears relieved;
How precious did that grace appear
The hour I first believed!

SERMON (Jane Smith)
```

---

## 📦 Building for Distribution

To build the application for distribution using PyInstaller:

```bash
pyinstaller --onefile --windowed --icon=icon.ico --add-data "Central Mega Template v1.pptx;." --add-data "icon.ico;." --add-data "logo.png;." main.py
```

---

## 📋 Dependencies

- Python 3.x
- `tkinter`
- `python-pptx`
- `poetry`

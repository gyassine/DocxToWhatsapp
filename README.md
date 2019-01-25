# DocxToWhatsapp
[![made-with-python](https://img.shields.io/badge/Made%20with-Python-1f425f.svg)](https://www.python.org/) 
[![GPLv3 license](https://img.shields.io/badge/License-GPLv3-blue.svg)](https://choosealicense.com/licenses/gpl-3.0/)
[![Maintenance](https://img.shields.io/maintenance/no/2018.svg)](https://github.com/gyassine/DocxToWhatsapp/graphs/commit-activity)
[![GitHub version](https://img.shields.io/github/release/gyassine/DocxToWhatsapp.svg)](https://github.com/gyassine/DocxToWhatsapp/releases)

A small python script that will open a `.docx` file and do some modifications that will allow me to copy &amp; paste in Whatsapp.

![Screenshot](screenshot.png?raw=true "Screenshot")

### Prerequisites :interrobang:
You need :
- python 3.6
- pyinstaller (if you want to make an exe)
  - `pip install pyinstaller` 
  
### Getting Started
- Download the latest release
- Unzip and execute it on `Example.docx`
- Paste it into whatever you want [WhatsApp](https://web.whatsapp.com) or anything else
![What it does...](Description.png?raw=true "What it does...")

### Release History :notebook:
Actual version : 1.1
####  V1.0
- New version modified to allow pyinstaller create an `.exe` file from `.py` file 
####  V0.2
- Implementation of GUI and "copy&paste" feature
####  V0.1
- Initial Commit Basic

## TODO :boom: 
- [X] :white_check_mark: GUI 
- [X] :clipboard: Copy & Paste
- [X] :newspaper: Update all my README.md to add nice badges
- [ ] :bookmark_tabs: Create a pdf file
- [ ] :inbox_tray: Upload PDF to webpage and add link

## Dependency :link:

- python : [pyperclip](https://pypi.python.org/pypi/pyperclip)
- Create an executable with [pyinstaller](https://www.pyinstaller.org/)
  - `pyinstaller .\DocxToWhatsapp.py --clean -w --icon=AdaboulMufrad.ico --noconfirm`

## Author :computer:	
- **Yassine Gangat** - [gyassine](https://github.com/gyassine)

## Contributing :ok_hand:
1. Fork it (<https://github.com/gyassine/DocxToWhatsapp/fork>)
2. Create your feature branch (`git checkout -b feature/fooBar`)
3. Commit your changes (`git commit -am 'Add some fooBar'`)
4. Push to the branch (`git push origin feature/fooBar`)
5. Create a new Pull Request

## Acknowledgments :thumbsup:
* Inspired by <https://gist.github.com/etienned/7539105> which only extract paragraph from word file

## License :scroll:

The content of this project itself is licensed under the [Creative Commons Attribution 3.0 license](http://creativecommons.org/licenses/by/3.0/us/deed.en_US), and the underlying source code used to format and display that content is licensed under the [GNU General Public License v3.0](https://choosealicense.com/licenses/gpl-3.0/).

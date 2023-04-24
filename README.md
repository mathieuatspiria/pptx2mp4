# pptx2mp4

pptx2mp4 converts a `.pptx` file to a video, and read out speaker's notes along slides using text to speech technique.

## Usage

* To get started, clone this repo
```
git clone https://github.com/mvillegas13/pptx2mp4
cd pptx2mp4
```
* Install python 3.10 (other versions will not be able to install the required TTS package at the time of writing). If you are on macOS, it is highly recommended to use homebrew to install all required software in this documentation (python, ffmpeg and libreoffice). To install python with homebrew, run the following command:
```
brew install python@3.10
```

* If you are on macOS, with Apple ARM cpu you need to install mecab too:
```
brew install mecab
```

* Install required packages
```
pip install --use-pep517 -r requirements.txt
```
on macOs use this command:
```
pip3.10 install --use-pep517 -r requirements.txt
```
* You also need:
    * [`ffmpeg`](https://github.com/adaptlearning/adapt_authoring/wiki/Installing-FFmpeg)
    * [`libreoffice`](https://www.libreoffice.org/download/download/)
    * Both 'ffmpeg' and 'soffice' software should be in your PATH

    If you are un macos just run:
    ```
    brew install ffmpeg libreoffice
    ```

* Run the script using the following command:
```
python pptx2mp4.py -i <input_pptx_file> -l <language_fr_or_en> -o <output_mp4_file>
```
on macOs use this command:
```
python3.10 pptx2mp4.py -i <input_pptx_file> -l <language_fr_or_en> -o <output_mp4_file>
```

To force the use of gtts (google text to speech) use the option "-t gtts"


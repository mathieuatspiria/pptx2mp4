# pptx2mp4

pptx2mp4 converts a `.pptx` file to a video, and read out speaker's notes along slides using text to speech technique.

## Usage

* To get started, clone this repo
```
git clone https://github.com/mvillegas13/pptx2mp4
cd pptx2mp4
```
* Install required packages
```
pip install -r requirements.txt
```
* You also need:
    * [`ffmpeg`](https://github.com/adaptlearning/adapt_authoring/wiki/Installing-FFmpeg)
    * [`libreoffice`](https://www.libreoffice.org/download/download/)
    * Both 'ffmpeg' and 'soffice' software should be in your PATH

* Run the script using the following command:
```
python pptx2mp4.py -i <input_pptx_file> -l <language_like_fr_en> -o <output_mp4_file>
```


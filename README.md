# didi_hackathon_pptx_creator
Simple script to create a Powerpoint slideshow with images.

Usage:
------
The script can be executed from the command line.

```
python pptx_create.py [-h] [-v] [-o OUTPUT_FILE_NAME] first_dir second_dir
```

Arguments:
```
Create a PowerPoint presentation with two images per slide.

positional arguments:
  first_dir             Directory containing the first set of images
  second_dir            Directory containing the second set of images

options:
  -h, --help            show this help message and exit
  -v, --verbose         Enable verbose mode
  -o OUTPUT_FILE_NAME, --output_file_name OUTPUT_FILE_NAME
                        Name of the output PowerPoint presentation
```

Requirements:
`python-pptx` and `pillow`. Can be installed with:

```
pip install -r requirements.txt
```


Example
-------
This repo contains a folder `test` with two subfolders `first_dir` and `second_dir` with two sets of seven images.

By running the command 

```
python pptx_create.py -v  test/first_dir test/second_dir
```

or on windows

```
python.exe pptx_create.py -v  tes\first_dir test\second_dir
```

You'll get the `output.pptx` also saved in this repo.

The images are scaled as such, that they fill the whole slide space while maintaining aspect ratio.



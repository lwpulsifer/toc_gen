# A table of contents generator for Google Slides with Google Slides API

## Overview

This is a quick little CLI to automate the boring and tedious
project of creating Table of Contents slides, made with Real Python video contributers in mind (I find myself making a whole lot of these slides). It works by copying the Real Python template slide from a presentation whose ID you must pass in. 
The ID of a Google Slides presentation can be found in the middle of the web address of the presentation like so:
Googledocs...com/presentation/d/**1o6u_OiH...zdDHCQ_iWrcv22mFybMOuE**/edit#slide=id.g730a0d59c0_0_0

The ID is the bold section in the middle of the URL.

The Python script then populates the table of contents in correct Real Python formatting using the data from a YAML config file.

## Usage

First, copy the Google Slides presentation at [this link](https://docs.google.com/presentation/d/1kDXfk_NV_dPJXezYFngJqCOs63T2RqR_8xtwpNF5Xag/edit?usp=sharing) into your own drive. Then copy the id into the `base_id` field of a yaml config file like `toc_example.yaml`, then run `toc.py` with that yaml file as the argument. The yaml should follow the format of the example exactly.

```bash
python toc.py CONFIG_FILE.yaml
```
When you run this, you'll prompted to log into your Google account and allow permissions, which it will need to work. You may also need to work with Google Developer tools to make sure this all works out, I'm not sure. I'm still learning all of this stuff.


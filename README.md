# pptx_audio


The code in this repository allows you to take Power Point and add audio or video to it, for teaching purposes. Here's how it works:

## Prerequisites

- A gpt-4o model deloyed to Azure Open AI
- A speech resource deployed to Azure
- Keys and config stored in .env file (see .env.example)
- Python runtime installed

## Preparation (manual step)

1. Open any PowerPoint deck and for each slide, delete the notes section
1. Update every slide, to include a prompt in the notes section on how the virtual teacher should explain the current slide
1. You can refer to any public URL, as the tool will use the content for grounding the slide transcript

## Run the tool (automated)

1. Run either video.py (slow) or audio.py (faster) with the pptx as the input.
1. The code will loop over all the slides in the deck (this might take time)
1. For each slide, if there's anything in the notes, it will send this as a prompt via langchain to Azure Open AI (gpt-4o model should be deployed). If a URL is present, it will fetch the content of the URL and transform to markdown (grounding the prompt)
1. In case of audio.py, it will transform the transcript obtain the previous step, to SSML and send this to the text to speech API for retrieving the mp3
1. In case of video.py, it will send the transcript to the video avatar endpoint and wait for the mp4.
1. mp3 or mp4 will be added the slide. Unfortunatly, there is no option for auto_play today (limitation of the ppxt library used) 
1. The output pptx will be saved in the pptx folder with the name: name_output.pptx.

```sh
python -m venv .venv
pip install -r requirements.txt
```

## Limitations

- There is no auto play for video/audio. For the audio: it's located at the top left corner of every slide (if present) and not visible on screen, unless you hover over it with your mouse/pointer.
- Video's do not have a transparant background (although it should be supported, but can't get that correctly working for now) - PowerPoint does not support webm, which seems to be required for transparancy.
- It takes a lot of time to transform, especially for video. Would like to consider working in parallel and not having to wait (async pattern)
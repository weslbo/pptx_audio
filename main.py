import sys
import getpass
import os
import bs4
import markdownify
import re
import azure.cognitiveservices.speech as speechsdk
from pptx import Presentation
from readability import Document
from bs4 import BeautifulSoup
from dotenv import load_dotenv
from langchain import hub
from langchain_openai import AzureChatOpenAI
from langchain_core.tools import tool
from langchain_core.messages import HumanMessage, SystemMessage
from langchain_unstructured import UnstructuredLoader
from langchain_community.document_loaders import WebBaseLoader
from langchain_core.prompts import ChatPromptTemplate
from langchain.agents import (AgentExecutor, create_tool_calling_agent)

# Define a very simple tool function that returns the current time
@tool
def get_current_time(*args, **kwargs):
    """Returns the current time in H:MM AM/PM format."""
    import datetime  # Import datetime module to get current time

    now = datetime.datetime.now()  # Get current time
    return now.strftime("%I:%M %p")  # Format time in H:MM AM/PM format

@tool
def retrieve_html(url):
    """Returns the markdown of the webpage at the given URL."""
    import requests
    response = requests.get(url)

    doc = Document(response.content)

    markdown = markdownify.markdownify(str(doc.summary()), heading_style="ATX")
    markdown = re.sub('\n{3,}', '\n\n', markdown)
    markdown = markdown.replace("[Continue](/en-us/)", "")

    return markdown

def generate_text_to_speech(input, savelocation):
    """Tranforms text to speech audio."""
    print(f"- Creating audio {savelocation}")
    
    service_region = os.getenv("SPEECH_REGION")
    speech_key = os.getenv("SPEECH_API_KEY")
    speech_config = speechsdk.SpeechConfig(subscription=speech_key, region=service_region)
    speech_config.set_speech_synthesis_output_format(speechsdk.SpeechSynthesisOutputFormat.Audio24Khz96KBitRateMonoMp3)  

    file_config = speechsdk.audio.AudioOutputConfig(filename=savelocation)
    speech_synthesizer = speechsdk.SpeechSynthesizer(speech_config=speech_config, audio_config=file_config)  

    result = speech_synthesizer.speak_text_async(input).get()
    return result

def main():
    load_dotenv()

    llm = AzureChatOpenAI(
        azure_endpoint=os.environ["AZURE_OPENAI_ENDPOINT"],
        azure_deployment=os.environ["AZURE_OPENAI_DEPLOYMENT_NAME"],
        openai_api_version=os.environ["AZURE_OPENAI_API_VERSION"],
    )

    tools = [retrieve_html]

    prompt = ChatPromptTemplate.from_messages(
        [
            (
                "system",
                """INSTRUCTION: Discuss the below input, following these guidelines:
                - You are a technical instructor. Your goal is to explain the topics in a clear and concise manner.
                - Use simple language and avoid jargon.
                - Only provide information that can be found in the content. Do not talk about anything that has not been mentioned in the prompt context.
                - Never output markdown syntax, code fragments, bulleted lists, etc. Remember, you are speaking, not writing.
                - Use conversational language to present the text's information.
                - You don't have to introduce yourself, greet listeners, or say goodbye. Jumpt straight into explaining the topic.
                - Add Variations in rhythm, stress, and intonation of speech depending on the context and statement.
                - Sometimes use filler words such as um, uh, you know and some stuttering but do not exaggerate please.
                - Do not use the word "Alright" to start the conversation.
                - You don't have to thank listeners or ask for questions. 
                - End the conversation abruptly after having discussed all the topics.""",
            ),
            ("human", "{input}"),
            ("placeholder", "{agent_scratchpad}")
        ]
    )
    
    agent = create_tool_calling_agent(llm, tools, prompt)
    agent_executor = AgentExecutor(agent=agent, tools=tools, verbose=True)

    pptx_input = "pathtopptxfile"
    pptx_output = pptx_input.replace("-notes.pptx", "-audio.pptx")

    presentation = Presentation(pptx_input)
    for slide in presentation.slides:
        if slide.notes_slide is not None:
            notes = slide.notes_slide.notes_text_frame.text

            if notes:
                response = agent_executor.invoke({"input": notes})
                output = response['output']
                slide.notes_slide.notes_text_frame.text = output
                print("response:", output)
                

                # Save the audio
                generate_text_to_speech(output, "output.mp3")
                audio = slide.shapes.add_movie("output.mp3", 0, 0, 1, 1)

    presentation.save(pptx_output)

if __name__ == "__main__":
    main()
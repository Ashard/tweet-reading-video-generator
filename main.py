import shutil

import PySimpleGUI as sg
import openpyxl
import requests
import os
import json
import demoji
import re
from moviepy.audio.io.AudioFileClip import AudioFileClip
from moviepy.audio.AudioClip import CompositeAudioClip
from moviepy.video.VideoClip import ImageClip
from moviepy.video.io.VideoFileClip import VideoFileClip
from moviepy.video.compositing.concatenate import concatenate_videoclips
from moviepy.audio.AudioClip import concatenate_audioclips
from moviepy.audio.fx.volumex import volumex
from moviepy.video.fx.resize import resize
from moviepy.video.fx.crop import crop
from moviepy.audio.fx.audio_fadeout import audio_fadeout

# Constants
EXECUTABLE_PATH = os.path.dirname(__file__)
IMAGE_EXTENSIONS = [".jpg", ".png"]
VIDEO_EXTENSIONS = [".mp4", ".avi"]
BG_AUDIO_FADEOUT = 3 # seconds

# main logic
def create_video(values):
    excel_file_path = values.get(EXCEL_FILE)
    wb_obj = openpyxl.load_workbook(excel_file_path)
    sheet = wb_obj.active

    tweet_img_paths = []
    tweet_audio_paths = []
    for row in sheet.iter_rows():
        for cell in row:
            tweet_link = cell.value
            if tweet_link == None:
                continue

            tweet_link = os.path.basename(tweet_link)
            print("\nGenerating tweet image for tweetID: {}".format(tweet_link))
            tweetImgPath, text = download_tweet(tweet_link)
            text = demoji.replace(text) # remove emojis
            text = re.sub(r"http\S+", "", text) # remove any http links

            if tweetImgPath != None:
                print("Tweet image successfully generated for tweetID: {}".format(tweet_link))
                tweet_img_paths.append(tweetImgPath)
            else:
                print("Tweet image could not be generated for tweetID: {}".format(tweet_link))

            print("\nGenerating TTS for tweetID: {}".format(tweet_link))
            voice_id = str(values.get(VOICE))
            reading_speed = str(values.get(READING_SPEED))
            reading_volume = str(values.get(READER_VOLUME))
            reading_pitch = str(values.get(READER_PITCH))

            tweetAudioPath = download_tts(tweet_link, text, voice_id=voice_id, speed=reading_speed,
                                          volume=reading_volume, pitch=reading_pitch)
            if tweetAudioPath != None:
                print("Tweet audio successfully generated for tweetID: {}".format(tweet_link))
                tweet_audio_paths.append(tweetAudioPath)
            else:
                print("Tweet audio could not be generated for tweetID: {}".format(tweet_link))

    if len(tweet_img_paths) != len(tweet_audio_paths):
        print("Error: Number of tweet images and audio files are not the same")
        return None

    tweet_img_paths = ['E:\\python projects\\ladyfawk\\res\\tweet_image_1484609415982448643.png', 'E:\\python projects\\ladyfawk\\res\\tweet_image_1484653157502402560.png', 'E:\\python projects\\ladyfawk\\res\\tweet_image_1484653703609167877.png']
    tweet_audio_paths = ['E:\\python projects\\ladyfawk\\res\\tweet_audio_1484609415982448643.mp3', 'E:\\python projects\\ladyfawk\\res\\tweet_audio_1484653157502402560.mp3', 'E:\\python projects\\ladyfawk\\res\\tweet_audio_1484653703609167877.mp3']

    tweet_image_clips = []
    tweet_audio_clips = []

    delay = 1
    silent_audio_clip = AudioFileClip(os.path.join(EXECUTABLE_PATH, "silence.mp3"))
    silent_audio_clip = silent_audio_clip.set_duration(delay)

    resolution = (1920, 1080)
    aspect_ratio = values.get(ASPECT_RATIO)
    if aspect_ratio != "16:9":
        if aspect_ratio == "1:1":
            resolution = (1080, 1080)
        elif aspect_ratio == "9:16":
            resolution = (1080, 1920)

    i = 0
    while i < len(tweet_img_paths):
        audio_clip = AudioFileClip(tweet_audio_paths[i])
        audio_dur = audio_clip.duration
        tweet_audio_clips.append(audio_clip)
        tweet_audio_clips.append(silent_audio_clip)

        tweet_image_clip = ImageClip(tweet_img_paths[i])
        tweet_image_clip = crop(tweet_image_clip, y1=120)
        tweet_image_clip = tweet_image_clip.set_duration(audio_dur + delay)
        tweet_image_clip = resize(tweet_image_clip, resolution)
        tweet_image_clips.append(tweet_image_clip)

        i += 1

    tweet_reading_video = concatenate_videoclips(tweet_image_clips)
    tweet_section_dur = tweet_reading_video.duration

    tweet_reading_audio = concatenate_audioclips(tweet_audio_clips)

    tweet_reading_video.audio = tweet_reading_audio
    tweet_reading_video = tweet_reading_video.set_duration(tweet_section_dur)

    intro_clip = None
    intro_clip_ext = os.path.splitext(values.get(INTRO_PATH))[1]
    if intro_clip_ext in IMAGE_EXTENSIONS:
        intro_clip = ImageClip(values.get(INTRO_PATH))
        intro_dur = int(values.get(INTRO_DURATION))
        intro_clip = intro_clip.set_duration(intro_dur)
    elif intro_clip_ext in VIDEO_EXTENSIONS:
        intro_clip = VideoFileClip(values.get(INTRO_PATH))
    intro_clip = resize(intro_clip, resolution)

    outro_clip = None
    outro_clip_ext = os.path.splitext(values.get(OUTRO_PATH))[1]
    if outro_clip_ext in IMAGE_EXTENSIONS:
        outro_clip = ImageClip(values.get(OUTRO_PATH))
        outro_dur = int(values.get(OUTRO_DURATION))
        outro_clip = outro_clip.set_duration(outro_dur)
    elif outro_clip_ext in VIDEO_EXTENSIONS:
        outro_clip = VideoFileClip(values.get(OUTRO_PATH))
    outro_clip = resize(outro_clip, resolution)

    bg_music_volume = int(values.get(BG_MUSIC_VOLUME))/100
    bg_music_clip = AudioFileClip(values.get(BG_MUSIC))
    bg_music_clip = volumex(bg_music_clip, bg_music_volume)

    actual_vid_duration = intro_clip.duration + outro_clip.duration + tweet_section_dur
    final_video = concatenate_videoclips([intro_clip, tweet_reading_video, outro_clip], method="compose")
    final_audio = CompositeAudioClip([final_video.audio, bg_music_clip])
    final_audio = final_audio.set_duration(actual_vid_duration)
    final_audio = audio_fadeout(final_audio, BG_AUDIO_FADEOUT)
    final_video.audio = final_audio

    final_video = final_video.set_duration(actual_vid_duration)
    output_path = os.path.join(EXECUTABLE_PATH, values.get(OUTPUT_FILENAME) + ".mp4")
    final_video.write_videofile(preset="ultrafast", threads=os.cpu_count()*4,
                                filename=output_path, audio_codec="aac",
                                fps=24)

    intro_clip.close()
    outro_clip.close()
    wb_obj.close()

    for f in os.listdir("res"):
        file_path = os.path.join("res", f)
        os.unlink(file_path)

# Downloads the text to speechaudio for the input text,
# and returns the path to it
def download_tts(tweetId, text, engine="neural",
                 voice_id="ai1-Matthew", language_code="en-US",
                 speed="0", volume="0", pitch="0"):
    headers = {
        'Content-Type': 'application/json',
        'Authorization': 'Bearer eedc47e0-779c-11ec-9a35-c3850f49dfad',
    }

    data = {
        "Text": "{}".format(text),
        "Engine": "{}".format(engine),
        "VoiceId": "{}".format(voice_id),
        "LanguageCode": "{}".format(language_code),
        "OutputFormat": "mp3",
        "SampleRate": "48000",
        "Effect": "default",
        "MasterSpeed": "{}".format(speed),
        "MasterVolume": "{}".format(volume),
        "MasterPitch": "{}".format(pitch)
    }
    data = json.dumps(data)
    api_url = "https://developer.voicemaker.in/voice/api"
    response = requests.post(api_url, headers=headers, data=data)

    if response.ok:
        response_json = response.json()
        download_url = response_json["path"]

        audio_response = requests.get(download_url)
        audio_path = os.path.join(EXECUTABLE_PATH, "res", "tweet_audio_{}.mp3".format(tweetId))
        audio_file = open(audio_path, "wb")
        audio_file.write(audio_response.content)
        audio_file.close()
        return audio_path
    return None

# Downloads the tweet as an image,
# then returns the path to the image and the text in the tweet as tuple
def download_tweet(tweetId):
    headers = {
        'Content-Type': 'application/json',
        'Authorization': '239e6047-914e-4fb9-975f-035b81d9c287',
    }

    tweetData = '"tweetId":"{}"'.format(tweetId)
    data = '{}{}{}'.format("{", tweetData, "}")
    response = requests.post('https://tweetpik.com/api/images', headers=headers, data=data)
    if response.ok:
        response_json = response.json()
        image_url = (response_json["url"])
        image_response = requests.get(image_url)
        image_name = "tweet_image_{}.png".format(tweetId)
        image_path = os.path.join(EXECUTABLE_PATH,"res", image_name)
        image_file = open(image_path, "wb")
        image_file.write(image_response.content)
        tweet_text = response_json["tweet"]["text"]
        image_file.close()
        return image_path, tweet_text
    return None


# =====================================UI CODE STARTS HERE ============================================
# Constants
INTRO_PATH =  "intro"
OUTRO_PATH = "outro"
BG_MUSIC = "background_music"
EXCEL_FILE = "excel_file"
READING_SPEED = "reading_speed"
INTRO_DURATION = "intro_duration"
OUTRO_DURATION = "outro_duration"
VOICE = "speaker"
READER_OPTIONS = ["ai1-Ivy", "ai1-Joanna", "ai1-Kendra", "ai1-Kimberly", "ai1-Salli", "ai1-Joey",
                   "ai1-Justin", "ai1-Kevin", "ai1-Matthew", "ai2-John", "ai2-Stacy", "ai2-Nikola",
                   "ai2-Scott", "ai2-Katie", "ai2-Scarlet", "ai2-Kathy", "ai2-Isabella", "ai2-Robert",
                   "ai2-Jerry", "ai3-Jony", "ai3-Aria", "ai3-Jenny", "ai3-Nova", "ai3-Olive", "ai3-Taylor",
                   "ai3-Vienna", "ai3-Kailey", "ai3-Addyson", "ai3-Emily", "ai3-Gary", "ai3-Kingsley", "ai3-Jason",
                   "ai3-Evan", "ai4-Sophia", "ai4-Amanda", "ai4-Edward", "ai4-Ronald", "ai4-Samantha",
                   "ai4-Roger", "ai4-Doris"]
READER_VOLUME = "reader_volume"
READER_PITCH = "reader_pitch"
OUTPUT_FILENAME = "output_filename"
BG_MUSIC_VOLUME = "background_music_volume"
ASPECT_RATIO = "aspect_ratio"

sg.theme("Dark Amber")
layout = [ [sg.Text(text="Intro"), sg.FileBrowse(key=INTRO_PATH)],
           [sg.Text(text="Intro duration (s)"), sg.InputText(key=INTRO_DURATION, default_text="3")],
           [sg.Text(text="Outro"), sg.FileBrowse(key=OUTRO_PATH)],
           [sg.Text(text="Outro duration (s)"), sg.InputText(key=OUTRO_DURATION, default_text="3")],
           [sg.Text(text="Background music"), sg.FileBrowse(key=BG_MUSIC)],
           [sg.Text(text="Background music volume"), sg.Slider(orientation="horizontal", key=BG_MUSIC_VOLUME, range=(0,100), default_value=20)],
           [sg.Text(text="Excel file"), sg.FileBrowse(key=EXCEL_FILE)],
           [sg.Text(text="Voice"), sg.DropDown(key=VOICE, values=READER_OPTIONS, default_value="ai1-Matthew")],
           [sg.Text(text="Reading speed"), sg.Slider(orientation="horizontal", key=READING_SPEED, default_value=0, range=(-100, 100))],
           [sg.Text(text="Reading volume"), sg.Slider(orientation="horizontal", key=READER_VOLUME, default_value=0, range=(-20, 20))],
           [sg.Text(text="Reading pitch"), sg.Slider(orientation="horizontal", key=READER_PITCH, default_value=0, range=(-100, 100))],
           [sg.Text(text="Aspect ratio"), sg.DropDown(key=ASPECT_RATIO, values=["1:1", "9:16", "16:9"], default_value="16:9")],
           [sg.Text(text="Output filename (exclude .mp4, etc)"), sg.InputText(key=OUTPUT_FILENAME, default_text="output")],
           [sg.Button(button_text="Submit")]
           ]

window = sg.Window("ladyfawk", layout)

def startUI():
    while True:
        event, values = window.read()
        if event == sg.WIN_CLOSED:
            break
        elif event == "Submit":
            create_video(values)

startUI()
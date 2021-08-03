import speech_recognition as sr
from os import path

from pydub import AudioSegment

sound = AudioSegment.from_mp3('1.mp3')
sound.export('output.wav', format="wav")
# transcribe audio file
# subprocess.call(['ffmpeg', '-i', '1.mp3',
#                  'audio.wav'])
AUDIO_FILE = "output.wav"

# use the audio file as the audio source

r = sr.Recognizer()

with sr.AudioFile(AUDIO_FILE) as source:

    audio = r.record(source) # read the entire audio file

print("Transcription: " + r.recognize_google(audio, language='es-ES'))
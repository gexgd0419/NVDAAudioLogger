# NVDA Addon: Audio Logger

This addon records system audio output and WavePlayer output, then saves them to a ZIP file, for debugging purposes.

## Usage

Press NVDA+Alt+R to start recording.

Press NVDA+Alt+T to stop recording and save the file.

Although you can leave it recording for an arbitrary duration, only the last 60 seconds of audio will be preserved. If you are trying to reproduce an issue, make sure to stop recording before the recorded issue goes pass the last 60 seconds.

The ZIP file will be saved in a folder named `AudioLogger` in your Documents folder. The file name is the date and time when the file was saved.

You can then send the ZIP file to the developers for diagnosis. The ZIP file is usually small enough to be able to be sent directly on GitHub.

## What is in the ZIP file

The ZIP file contains several WAV audio files.

`SystemAudio.wav` is the recorded system audio output, including audio emitted by other apps.

`WavePlayer X.wav` contains the original wave data sent to a WASAPI WavePlayer, usually the original audio data from a voice.

Asides from audio, gesture inputs (such as keyboard inputs) and spoken texts will also be included in those WAV files, in the form of markers. You can use some audio editing software to check the markers.

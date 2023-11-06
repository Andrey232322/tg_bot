from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
import telebot
import pythoncom
import os
from datetime import datetime,timedelta
import screen_brightness_control as sbc
bot = telebot.TeleBot('Token_bot')

def up(num):

    num = int(num) / 100
    devices = AudioUtilities.GetSpeakers()
    interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
    volume = cast(interface, POINTER(IAudioEndpointVolume))

    current_volume = round(volume.GetMasterVolumeLevelScalar(),2)
    volume.SetMasterVolumeLevelScalar(current_volume, None)
    vol = volume.GetMasterVolumeLevelScalar() + num if volume.GetMasterVolumeLevelScalar() + num < 0.87 else current_volume
    return volume.SetMasterVolumeLevelScalar(vol ,None)

def down(num):

    num = int(num) / 100
    devices = AudioUtilities.GetSpeakers()
    interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
    volume = cast(interface, POINTER(IAudioEndpointVolume))

    current_volume = volume.GetMasterVolumeLevelScalar()
    volume.SetMasterVolumeLevelScalar(current_volume, None)
    vol = volume.GetMasterVolumeLevelScalar() - num if volume.GetMasterVolumeLevelScalar() - num > 0 else current_volume
    return volume.SetMasterVolumeLevelScalar(vol, None)



def time_off(house):
    a = datetime.now()
    b = a + timedelta(minutes=house)
    while a != b:
        a = datetime.now()
        if a.strftime('%H:%M') == b.strftime('%H:%M'):
            return os.system('shutdown -s')

def fade(int):
    return sbc.set_brightness(int)

@bot.message_handler(commands=['up'])
def add_volume(message):
    pythoncom.CoInitializeEx(0)
    count = int(message.text.replace('/up', '').strip()) if message.text else None
    return up(count)


@bot.message_handler(commands=['down'])
def down_volume(message):
    pythoncom.CoInitializeEx(0)
    count = int(message.text.replace('/down', '').strip()) if message.text else None
    return down(count)


@bot.message_handler(commands=['time'])
def down_volume(message):
    count = int(message.text.replace('/time', '').strip()) if message.text else None
    return time_off(count)

@bot.message_handler(commands=['off'])
def down_volume(message):
    pythoncom.CoInitializeEx(0)
    return off()


@bot.message_handler(commands=['fade'])
def down_volume(message):
    pythoncom.CoInitializeEx(0)
    count = int(message.text.replace('/fade', '').strip()) if message.text else None
    return fade(count)


@bot.message_handler(commands=['help'])
def down_volume(message):
    return bot.send_message(message.from_user.id, "up-прибавить звук,  "
                                                  "down-убавить звук,   "
                                                  "off-выключить комп,  "
                                                  "time- выключить комп через время в минутах,  "
                                                  "fade- яркость");
def off():
    return os.system('shutdown -s')


bot.polling(none_stop=True, interval=0)
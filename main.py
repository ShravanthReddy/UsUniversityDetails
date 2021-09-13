import os
import telebot
from telebot import apihelper
import openpyxl
from settings import API_KEY

bot = telebot.TeleBot(API_KEY)
apihelper.SESSION_TIME_TO_LIVE = 60 * 5

@bot.message_handler(commands=['start'])
def start(message):
  bot.send_message(message.chat.id,'Hello, welcome!\nHere, you can find Application & Tution fee details of various Universities.\nTap /universitysearch to begin')

@bot.message_handler(commands=['universitysearch'])
def uniname(message):
  msg = bot.send_message(message.chat.id,'Enter the University name or city')
  bot.register_next_step_handler(msg, comparison)

def comparison(message):
  collegeByUser = message.text
  collegeByUserLower = collegeByUser.lower()
  wb = openpyxl.load_workbook('USUniDetails.xlsx')
  MainSheet = wb['Sheet1']
  CollegeName = list()
  count = 0

  for i in range(1, 329):
    CollegeName.append(MainSheet.cell(row=i,column=2).value)
    converted_list = [x.lower() for x in CollegeName]
    if any(collegeByUserLower in word for word in converted_list):
      CollegeNameFinal = ''.join(CollegeName)
      ApplicationFee = str(MainSheet.cell(row=i,column=3).value)
      TutionFee = str(MainSheet.cell(row=i,column=4).value)

      bot.send_message(message.chat.id, 'College Name: ' + CollegeNameFinal + '\nApplication Fee: ' + ApplicationFee + '\nTution Fee/Year: ' + TutionFee)
      count = count+1

    CollegeName = list()
  if count == 0:
    bot.send_message(message.chat.id, 'No details found')
    bot.send_message(message.chat.id, 'For more information, contact a consultancy near you or check the University website.\nTo search for another University, tap /universitysearch')
  else:
    bot.send_message(message.chat.id, 'This is an estimate and is not guaranteed. For more details, contact a consultancy near you or check the University website.\nTo search for another University, tap /universitysearch')
bot.polling()

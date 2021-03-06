import telebot
from telebot import apihelper, types
import openpyxl
from settings import API_KEY
import time as t

while True:
  #Bot Initialization
  bot = telebot.TeleBot(API_KEY)
  apihelper.SESSION_TIME_TO_LIVE = 60 * 5

  #Initializing Keyboard Markup buttons text
  optionA = 'Search University by name or City'
  optionB = 'Search Universities offering application fee waiver'
  optionC = 'Search by name or city'
  optionD = 'All Universities'
  optionE = 'Go Back'

  #Initializing Keyboard Markup buttons
  markup = types.ReplyKeyboardMarkup(row_width=1, one_time_keyboard=True)
  itembtn1 = types.KeyboardButton(optionA)
  itembtn2 = types.KeyboardButton(optionB)
  markup.add(itembtn1, itembtn2)

  markup1 = types.ReplyKeyboardMarkup(row_width=1, one_time_keyboard=True)
  itembtn3 = types.KeyboardButton(optionC)
  itembtn4 = types.KeyboardButton(optionD)
  itembtn5 = types.KeyboardButton(optionE)
  markup1.add(itembtn3, itembtn4, itembtn5)

  #Start of the program
  @bot.message_handler(commands=['start'])
  def start(message):
    reply = bot.send_message(message.chat.id, 'Hello, welcome!\nHere, you can find Application & Tution fee details of various Universities.\nChoose an option to continue: ', reply_markup=markup)
    bot.register_next_step_handler(reply, check)

  @bot.message_handler(commands=['continue'])
  def cont(message):
    reply = bot.send_message(message.chat.id,'Choose an option to continue: ', reply_markup=markup)
    bot.register_next_step_handler(reply, check)

  @bot.message_handler(func=lambda message: True)
  def all(message):
    bot.send_message(message.chat.id, 'Wrong input, please try again. \nTo search for university details, tap /continue or /start')

  #Checking the option selected in first page
  def check(message):
    if message.text == optionA:
      UniversitySearch(message)
    elif message.text == optionB:
      afwUni(message)
    else:
      bot.send_message(message.chat.id, 'Wrong option selected, please try again')
      cont(message)

  #Checking the option selected in second page    
  def check2(message):
    if message.text == optionC:
      reply2 = bot.send_message(message.chat.id, 'Please enter the University name or city')
      bot.register_next_step_handler(reply2, afwUniversitySearch)
    elif message.text == optionD:
      bot.send_message(message.chat.id, 'Showing results for all the Universities with Application fee waiver')
      afwAll(message)
    elif message.text == optionE:
      cont(message)
    else:
      bot.send_message(message.chat.id, 'Wrong option selected, please try again')
      afwUni(message)

  #Function to capture user input for university name or city
  def UniversitySearch(message):
    msg = bot.send_message(message.chat.id, 'Please enter the University name or city')
    bot.register_next_step_handler(msg, universitySearch)

  #Function which displays page two options
  def afwUni(message):
    reply2 = bot.send_message(message.chat.id, 'Choose one below: ', reply_markup=markup1)
    bot.register_next_step_handler(reply2, check2)

  #Function to capture user input and change it to lower text
  def constants(message):
    collegeByUser = message.text
    collegeByUserLower = collegeByUser.lower()
    return collegeByUserLower

  #Function to initialize excel file
  def excel():
    wb = openpyxl.load_workbook('USUniDetails.xlsx')
    MainSheet = wb['Sheet1']
    return MainSheet
  
  #String Split function
  def stringSplit(inputString):
    finalStringList = list()
    #print(inputString)
    splitInputString = inputString.split(',')
    spiltInputStringLen = len(splitInputString)
    #print(splitInputString)

    for i in range(0, spiltInputStringLen):
            splitInputStringWithSpaces = splitInputString[i].split()
            finalStringList = finalStringList + [x.lower () for x in splitInputStringWithSpaces]
    return finalStringList

  #University search by name or city
  def universitySearch(message):
    MainSheet = excel()
    collegeByUserLowerA = constants(message)
    collegeByUserLower = list()
    collegeByUserLower = stringSplit(collegeByUserLowerA)
    collegeByUserLowerLength = len(collegeByUserLower)
    collegeName = list()
    count = 0
    maxcount = 0
    final = 0
    collegeMatchList = {}
    applicationFeeList = {}
    tutionFeeList = {}
    for college in range(1, 329):
      collegeName = list()
      collegeName.append(MainSheet.cell(row=college, column=2).value)
      finalCollegeName = list()
      collegeNameString = ''.join(collegeName)
      finalCollegeName = stringSplit(collegeNameString)
      length2 = len(finalCollegeName)
      count = 0

      for i in range(0, collegeByUserLowerLength):
          for j in range(0, length2):
              #print('User Input Word: ', userInputList[i], end=" ")
              if finalCollegeName[j] != 'university' and finalCollegeName[j] != 'of' and finalCollegeName[j] != '(web)' and finalCollegeName[j] != 'at' and finalCollegeName[j] == collegeByUserLower[i]:
                  #print("Matched with ", finalCollegeName[i], end="")
                  count = count+1

      if count > 0:
          collegeMatchList[collegeNameString] = count 
          applicationFeeList[collegeNameString] = str(MainSheet.cell(row=college, column=3).value)
          tutionFeeList[collegeNameString] = str(MainSheet.cell(row=college, column=4).value)

      if count > maxcount:
          maxcount = count

    for collegeName, count in collegeMatchList.items():
      if count == maxcount:
          applicationFee = applicationFeeList.get(collegeName)
          tutionFee = tutionFeeList.get(collegeName)
          bot.send_message(message.chat.id, 'College Name: ' + collegeName + '\nApplication Fee: ' + applicationFee + '\nTution Fee/Year: ' + tutionFee)
          final = final+1

    endMessage(message, final)

  #AFW university search by name or city
  def afwUniversitySearch(message):
    MainSheet = excel()
    collegeByUserLowerA = constants(message)
    collegeByUserLower = list()
    collegeByUserLower = stringSplit(collegeByUserLowerA)
    collegeByUserLowerLength = len(collegeByUserLower)
    collegeName = list()
    count = 0
    maxcount = 0
    final = 0
    collegeMatchList = {}
    applicationFeeList = {}
    tutionFeeList = {}
    for college in range(1, 329):
      collegeName = list()
      collegeName.append(MainSheet.cell(row=college, column=2).value)
      finalCollegeName = list()
      collegeNameString = ''.join(collegeName)
      finalCollegeName = stringSplit(collegeNameString)
      length2 = len(finalCollegeName)
      count = 0
      ApplicationFee = str(MainSheet.cell(row=college, column=3).value)
      TutionFee = str(MainSheet.cell(row=college, column=4).value)
      if ApplicationFee == 'AFW' or ApplicationFee == 'Free':
        for i in range(0, collegeByUserLowerLength):
          for j in range(0, length2):
              #print('User Input Word: ', userInputList[i], end=" ")
              if finalCollegeName[j] != 'university' and finalCollegeName[j] != 'of' and finalCollegeName[j] != '(web)' and finalCollegeName[j] != 'at' and finalCollegeName[j] == collegeByUserLower[i]:
                  #print("Matched with ", finalCollegeName[i], end="")
                  count = count+1

      if count > 0:
          collegeMatchList[collegeNameString] = count 
          applicationFeeList[collegeNameString] = ApplicationFee
          tutionFeeList[collegeNameString] = TutionFee

      if count > maxcount:
          maxcount = count

    for collegeName, count in collegeMatchList.items():
      if count == maxcount:
          applicationFee = applicationFeeList.get(collegeName)
          tutionFee = tutionFeeList.get(collegeName)
          bot.send_message(message.chat.id, 'College Name: ' + collegeName + '\nApplication Fee: ' + applicationFee + '\nTution Fee/Year: ' + tutionFee)
          final = final+1

    endMessage(message, final)

  #AFW all universities
  def afwAll(message):
    MainSheet = excel()
    collegeName = list()
    count = 0
    for i in range(1, 329):
      collegeName.append(MainSheet.cell(row=i, column=2).value)
      ApplicationFee = str(MainSheet.cell(row=i, column=3).value)
      TutionFee = str(MainSheet.cell(row=i, column=4).value)
      if ApplicationFee == 'AFW' or ApplicationFee == 'Free':
        CollegeNameFinal = ''.join(collegeName)

        bot.send_message(message.chat.id, 'College Name: ' + CollegeNameFinal + '\nApplication Fee: ' + ApplicationFee + '\nTution Fee/Year: ' + TutionFee)
        count = count + 1
      collegeName = list()
    endMessage(message, count)

  #End message
  def endMessage(message, count):
    if count == 0:
      bot.send_message(message.chat.id, 'No details found')
      bot.send_message(message.chat.id, 'For more information, contact a consultancy near you or check the University website.\nTo continue searching for universities, tap /continue')
    else:
      bot.send_message(message.chat.id, 'This is an estimate and is not guaranteed. For more details, contact a consultancy near you or check the University website.\nTo continue searching for universities, tap /continue')
  try:
    bot.polling()
  except Exception as e:
    t.sleep(5)

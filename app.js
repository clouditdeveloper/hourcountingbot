require('dotenv').config(); // Загрузка переменных окружения из файла .env
const TelegramBot = require('node-telegram-bot-api');
const ExcelJS = require('exceljs');
const fs = require('fs');

// Создание экземпляра бота
const bot = new TelegramBot(process.env.TELEGRAM_BOT_TOKEN, { polling: true });

// Обработка команды /start
bot.onText(/\/start/, (msg) => {
  const chatId = msg.chat.id;

  const startMessage = `
  В мире времени, где минуты тают, я являюсь стражем секунд и мгновений — 🤖 бот учета рабочих часов. Я воплощаю ритм эффективности, внимательно следя за каждым прожитым мгновением, чтобы показать, что время, вложенное в труд, превращается в сокровище непреходящей ценности. ⌛✨
  `;

  const startButton = {
    reply_markup: {
      inline_keyboard: [
        [
          {
            text: 'Начать учет',
            callback_data: 'startTracking',
          },
        ],
      ],
    },
  };

  bot.sendMessage(chatId, startMessage, startButton);
});

// Обработка нажатия на кнопку "Начать учет"
bot.on('callback_query', (query) => {
  const chatId = query.message.chat.id;

  if (query.data === 'startTracking') {
    const date = getCurrentDate();
    const startTime = getCurrentTime();

    const workRecord = {
      date: date,
      startTime: startTime,
      breakTime: '45',
    };

    bot.sendMessage(chatId, `Дата: ${date}`);
    bot.sendMessage(chatId, `Время начала: ${startTime}`);
    bot.sendMessage(chatId, 'Вы закончили работу?', {
      reply_markup: {
        inline_keyboard: [
          [
            {
              text: 'Да',
              callback_data: 'endTracking',
            },
          ],
          [
            {
              text: 'Нет',
              callback_data: 'continueTracking',
            },
          ],
        ],
      },
    });

    bot.on('callback_query', (query) => {
      if (query.data === 'endTracking') {
        const endTime = getCurrentTime();
        workRecord.endTime = endTime;

        bot.sendMessage(chatId, `Время окончания: ${endTime}`);
        bot.sendMessage(chatId, 'Введите название объекта:');
        bot.once('message', (msg) => {
          const objectName = msg.text;
          workRecord.objectName = objectName;

          saveWorkRecord(chatId, workRecord);
          bot.sendMessage(chatId, 'Запись сохранена.');
        });
      } else if (query.data === 'continueTracking') {
        bot.sendMessage(chatId, 'Продолжайте работу.');
      }
    });
  }
});

// Обработка команды /stunden
bot.onText(/\/stunden/, (msg) => {
  const chatId = msg.chat.id;

  sendExcelFile(chatId);
});

// Функция сохранения записи о рабочих часах в таблицу Excel
function saveWorkRecord(chatId, workRecord) {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet('Рабочие часы');

  const filePath = `${chatId}.xlsx`;

  try {
    if (fs.existsSync(filePath)) {
      workbook.xlsx.readFile(filePath);
    } else {
      worksheet.addRow(['Datum', 'Von', 'Pause', 'Bis', 'Ort', '24st. Entf', 'Frei', 'Krank', 'Stunden']);
    }

    worksheet.addRow([
      workRecord.date,
      workRecord.startTime,
      workRecord.breakTime,
      workRecord.endTime,
      workRecord.objectName,
    ]);

    workbook.xlsx.writeFile(filePath);
  } catch (error) {
    console.error(error);
  }
}

// Функция отправки файла Excel
function sendExcelFile(chatId) {
  const filePath = `${chatId}.xlsx`;

  if (fs.existsSync(filePath)) {
    bot.sendDocument(chatId, filePath, { caption: 'Рабочие часы.xlsx' });
  } else {
    bot.sendMessage(chatId, 'Нет доступных записей о рабочих часах.');
  }
}

// Функция проверки корректности времени
function isValidTime(timeString) {
  const regex = /^([0-1]?[0-9]|2[0-3]):[0-5][0-9]$/;
  return regex.test(timeString);
}

// Функция получения текущей даты в формате ДД.ММ.ГГГГ
function getCurrentDate() {
  const currentDate = new Date();
  const day = String(currentDate.getDate()).padStart(2, '0');
  const month = String(currentDate.getMonth() + 1).padStart(2, '0');
  const year = currentDate.getFullYear();
  return `${day}.${month}.${year}`;
}

// Функция получения текущего времени в формате ЧЧ:ММ
function getCurrentTime() {
  const currentTime = new Date();
  const hours = String(currentTime.getHours()).padStart(2, '0');
  const minutes = String(currentTime.getMinutes()).padStart(2, '0');
  return `${hours}:${minutes}`;
}

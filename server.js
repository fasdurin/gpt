var excel = require('excel4node');
const mongoose = require('mongoose');
const express = require('express');
const app = express();
const path = require('path');
const bodyParser = require('body-parser');

app.use(bodyParser.json()); 
app.use(express.static(__dirname));

mongoose.connect('mongodb://localhost/duplom', { useNewUrlParser: true, useUnifiedTopology: true })
    .then(() => console.log('Підключено до MongoDB...'))
    .catch(err => console.error('Помилка підключення до MongoDB:', err));
    const disciplineSchema = new mongoose.Schema({
      surename: String,
      name: String,
      midle_name: String,
      year:Number,
      first_discepline: Object,
      second_discepline: Object,
      third_discepline: Object,
      four_discepline:Object,
      five_discepline:Object,
      six_discepline:Object
    }, { collection: 'VK_stydents' });
    const Discipline = mongoose.model('Discipline', disciplineSchema);
    const OK_disciplineSchema=new mongoose.Schema({
      conVK:String,
      name:String,
      countCredit:Number,
      course:Number,
      semester:Number,
      hoursInWeek:Number,
      countWeek:Number,
      formOfControl:String,
      lectures:Number,
      practicalLaboratory:Number,
      seminar:Number,
      year:Number
    },{collection:'OK_disceplines'});
    const OK_discepline=mongoose.model('OK_discepline',OK_disciplineSchema);
    const VK_disciplineSchema=new mongoose.Schema({
      conVK:String,
      name:String,
      countCredit:Number,
      course:Number,
      semester:Number,
      hoursInWeek:Number,
      countWeek:Number,
      formOfControl:String,
      lectures:Number,
      practicalLaboratory:Number,
      seminar:Number,
      year:Number
    },{collection:'VK_disciplines'});
    const VK_discepline=mongoose.model('VK_discepline',VK_disciplineSchema);
    const PracticeShema=new mongoose.Schema({
      conOK:String,
      name:String,
      countCredit:Number,
      course:Number,
      semester:Number,
      lenght:Number
    },{collections:'Practices'});
    const Practice=mongoose.model('Practice',PracticeShema);
    app.post('/api/students', (req, res) => {
      const studentData = req.body;
      console.log(studentData);
      addDiscipline(studentData);
      res.send('Дані успішно відправлено!');
    });
    app.post('/api/students_1',async (req,res) =>{
      const yearValue = req.body;
      const year= parseInt(yearValue.year);
      const data= await getVkdiscipline(year);
      res.json(data);
    });
    async function getVkdiscipline(tmp){
      let y=tmp-1;
      const studentVK= await VK_discepline.find({
        year:y
      });
      return studentVK;
    }
  app.post('/app/students', async (req, res) => {
    const student = req.body;
    console.log(student);
    const stydent = await getstudentData(student);
    getDisciplines(stydent);
    res.json(stydent); // Відправляємо дані назад на клієнт
  });
  async function getstudentData(tmp) {
    const stydent = await Discipline.find({
      surename: tmp.surename,
      name: tmp.name,
      midle_name: tmp.midlename,
    });
    console.log(stydent);
    return stydent; 
  }
    async function addDiscipline(studentData) {
      const discipline = new Discipline({
        surename: studentData.surename,
        name: studentData.name,
        midle_name: studentData.midle_name,
        year:studentData.year,
        first_discepline: studentData.first_discepline,
        second_discepline: studentData.second_discepline,
        third_discepline: studentData.third_discepline,
        four_discepline:studentData.four_discepline,
        five_discepline:studentData.five_discepline,
        six_discepline:studentData.six_discepline
      });
    
      const result = await discipline.save();
      console.log(result);
    }
  // Функція для отримання даних з колекції "OK disciplines" для 2 курсу 1 семестр
  async function getDisciplines(student) {
    let h=student[0].year-1;
    const Ok_disciplines = await OK_discepline.find({
      course:2,
      semester:1,
      year:h
    });
    const Ok_disciplines_2 = await OK_discepline.find({
      course:2,
      semester:2,
      year:h
    });
    const Ok_disciplines_3 = await OK_discepline.find({
      course:3,
      semester:3,
      year:h
    });
    const Ok_disciplines_4 = await OK_discepline.find({
      course:3,
      semester:4,
      year:h
    });
    const Ok_disciplines_5 = await OK_discepline.find({
      course:4,
      semester:5,
      year:h
    });
    const Ok_disciplines_6 = await OK_discepline.find({
      course:4,
      semester:6,
      year:h
    });
    const VK_S=[
      student[0].first_discepline.name,
      student[0].second_discepline.name,
      student[0].third_discepline.name,
      student[0].four_discepline.name,
      student[0].five_discepline.name,
      student[0].six_discepline.name
    ];
    const VK_disciplines=await VK_discepline.find({
      course:2,
      semester:1,
      name:{ $in: VK_S },
      year:h
    });
    const VK_disciplines_2=await VK_discepline.find({
      course:2,
      semester:2,
      name:{ $in: VK_S },
      year:h
    });
    const VK_disciplines_3=await VK_discepline.find({
      course:3,
      semester:3,
      name:{ $in: VK_S },
      year:h
    });
    const VK_disciplines_4=await VK_discepline.find({
      course:3,
      semester:4,
      name:{ $in: VK_S },
      year:h
    });
    const VK_disciplines_5=await VK_discepline.find({
      course:4,
      semester:5,
      name:{ $in: VK_S },
      year:h
    });
    const VK_disciplines_6=await VK_discepline.find({
      course:4,
      semester:6,
      name:{ $in: VK_S },
      year:h
    });

    const Practices= await Practice.find();

    // Створюємо новий екземпляр класу Workbook
    var workbook = new excel.Workbook();
    // Додаємо робочі аркуші до книги
    Title_page(workbook,student);
    var worksheet = workbook.addWorksheet('II курс I семестер');
    const Style0={
      font: {
        name: 'Times New Roman',
        size:12
      },
    };
  let i=student[0].year+1;
  worksheet.cell(1,1,1,3,true).string("Навчальний рік "+student[0].year+"/"+i).style(Style0);
  worksheet.cell(2,1,2,3,true).string("Курс: II").style(Style0);
  worksheet.cell(3,1,3,8,true).string("Семестр: I з_____"+student[0].year+"р."+ " до_____"+student[0].year+"р.").style(Style0);
  worksheet.cell(4,1,4,8,true).string("Екзаменаційна сесія з_____"+student[0].year+"р."+" до_____"+student[0].year+"р.").style(Style0);
  worksheet.cell(1,8,1,17,true).string("Прізвище, ім’я по батькові здобувача освіти: "+student[0].surename+" "+student[0].name+" "+student[0].midle_name).style(Style0);
  worksheet.cell(2,8,2,17,true).string("Група:               1-ІПЗ-"+(student[0].year-1)%100).style(Style0);
  worksheet.column(1).setWidth(5);
  worksheet.column(2).setWidth(5);
  worksheet.column(4).setWidth(5);
  worksheet.column(5).setWidth(5);
  worksheet.column(6).setWidth(5);
  worksheet.column(7).setWidth(5);
  worksheet.column(8).setWidth(5);
  worksheet.column(9).setWidth(5);
  worksheet.column(10).setWidth(5);
  worksheet.column(11).setWidth(5);
  worksheet.column(12).setWidth(5);
  
  const Style={
    alignment: {
      horizontal: 'center',
      vertical: 'center',
      wrapText: true,
    },
    font: {
      name: 'Times New Roman',
      bold: true,
      size:10
    },
    border: { 
      left: {
        style: 'thin',
        color: 'black'
      },
      right: {
        style: 'thin',
        color: 'black'
      },
      top: {
        style: 'thin',
        color: 'black'
      },
      bottom: {
        style: 'thin',
        color: 'black'
      }
    }
  };
  const Style_rotate={
    alignment:{
      textRotation: 90,
      horizontal: 'center',
      vertical: 'center',
      wrapText: true,
    },
    font: {
      name: 'Times New Roman',
      bold: true,
      size:10
    },
    border: { 
      left: {
        style: 'thin',
        color: 'black'
      },
      right: {
        style: 'thin',
        color: 'black'
      },
      top: {
        style: 'thin',
        color: 'black'
      },
      bottom: {
        style: 'thin',
        color: 'black'
      }
    }
  };
    worksheet.cell(6,1,8,1,true).string("№ з/п").style(Style);
    worksheet.cell(6,2,8,2,true).string("Код ОК").style(Style);
    worksheet.cell(6,3,8,3,true).string("Назва\n освітнього\n компонента").style(Style);
    worksheet.column(3).setWidth(20);
    worksheet.cell(6,4,8,4,true).string("Кількість кредитів ЄКТС").style(Style_rotate);
    worksheet.row(6).setHeight(20);
    worksheet.row(7).setHeight(70);
    worksheet.row(8).setHeight(90);
    worksheet.cell(6,5,6,11,true).string("Кількість годин").style(Style);
    worksheet.cell(7,5,8,5,true).string("Всього").style(Style_rotate);
    worksheet.cell(7,6,8,6,true).string("Аудиторні години").style(Style_rotate);
    worksheet.cell(7,7,7,11,true).string("З них").style(Style);
    worksheet.cell(8,7).string("Лекції").style(Style_rotate);
    worksheet.cell(8,8).string("Практичні,\nЛабораторні").style(Style_rotate);
    worksheet.cell(8,9).string("Семінарські\n заняття").style(Style_rotate);
    worksheet.cell(8,10).string("Самостійна\n робота").style(Style_rotate);
    worksheet.cell(8,11).string("Індивідуальна\n робота").style(Style_rotate);
    worksheet.cell(6,12,8,12,true).string("Форма підсумкового контролю").style(Style_rotate);
    worksheet.cell(6,13,7,14,true).string("Оцінка").style(Style);
    worksheet.cell(8,13).string("за 12- бальною\n шкалою").style(Style);
    worksheet.cell(8,14).string("за шкалою\n закладу освіти").style(Style);
    worksheet.column(13).setWidth(8);
    worksheet.column(14).setWidth(8);
    worksheet.cell(6,15,8,15,true).string("Дата\n проведення\n \nпідсумкового\n контролю").style(Style);
    worksheet.column(15).setWidth(11);
    worksheet.cell(6,16,8,16,true).string("Прізвище,\n ініціали\n викладача").style(Style);
    worksheet.column(16).setWidth(9);
    worksheet.cell(6,17,8,17,true).string("Підпис\n викладача").style(Style);
    worksheet.cell(9,1,9,17,true).string("Обов’язкові освітні компоненти").style(Style);
    worksheet.column(17).setWidth(9);
    const Style_={
      alignment:{
        horizontal: 'center',
        vertical: 'center',
        wrapText: true,
      },
      font: {
        name: 'Times New Roman',
        size:10
      },
      border: { 
        left: {
          style: 'thin',
          color: 'black'
        },
        right: {
          style: 'thin',
          color: 'black'
        },
        top: {
          style: 'thin',
          color: 'black'
        },
        bottom: {
          style: 'thin',
          color: 'black'
        }
      }
    };
    const Style_1={
      alignment:{
        vertical: 'center',
        wrapText: true,
      },
      font: {
        name: 'Times New Roman',
        size:10
      },
      border: { 
        left: {
          style: 'thin',
          color: 'black'
        },
        right: {
          style: 'thin',
          color: 'black'
        },
        top: {
          style: 'thin',
          color: 'black'
        },
        bottom: {
          style: 'thin',
          color: 'black'
        }
      }
    };
    let startRow = 10;
    let C=1;
    let startColumn = 1;
    Ok_disciplines.forEach(discipline=>{
      Object.keys(discipline).forEach((key,index)=>{
        worksheet.cell(startRow,startColumn).number(C).style(Style_);
        worksheet.cell(startRow,startColumn+1).string(discipline.conVK).style(Style_);
        worksheet.cell(startRow,startColumn+2).string(discipline.name).style(Style_1);
        worksheet.cell(startRow,startColumn+3).number(discipline.countCredit).style(Style_);
        worksheet.cell(startRow,startColumn+4).number(discipline.hoursInWeek*discipline.countWeek).style(Style_);
        worksheet.cell(startRow,startColumn+5).number(discipline.lectures+discipline.practicalLaboratory+discipline.seminar).style(Style_);
        worksheet.cell(startRow,startColumn+6).number(discipline.lectures).style(Style_);
        worksheet.cell(startRow,startColumn+7).number(discipline.practicalLaboratory).style(Style_);
        worksheet.cell(startRow,startColumn+8).number(discipline.seminar).style(Style_);
        worksheet.cell(startRow,startColumn+9).number(discipline.countCredit*30-discipline.hoursInWeek*discipline.countWeek).style(Style_);
        worksheet.cell(startRow,startColumn+11).string(discipline.formOfControl).style(Style_);
        worksheet.cell(startRow,startColumn+10).style(Style_);
        worksheet.cell(startRow,startColumn+12).style(Style_);
        worksheet.cell(startRow,startColumn+13).style(Style_);
        worksheet.cell(startRow,startColumn+14).style(Style_);
        worksheet.cell(startRow,startColumn+15).style(Style_);
        worksheet.cell(startRow,startColumn+16).style(Style_);
      });
      startRow++;
      C++;
    });
    const y=startRow-1;
    worksheet.cell(startRow,1).style(Style_);
    worksheet.cell(startRow,2).style(Style_);
    worksheet.cell(startRow,3).string("Всього:").style(Style_1);
    worksheet.cell(startRow,4).formula('SUM(D10:D'+y+')').style(Style_);
    worksheet.cell(startRow,5).formula('SUM(E10:E'+y+')').style(Style_);
    worksheet.cell(startRow,6).formula('SUM(F10:F'+y+')').style(Style_);
    worksheet.cell(startRow,7).formula('SUM(G10:G'+y+')').style(Style_);
    worksheet.cell(startRow,8).formula('SUM(H10:H'+y+')').style(Style_);
    worksheet.cell(startRow,9).formula('SUM(I10:I'+y+')').style(Style_);
    worksheet.cell(startRow,10).formula('SUM(J10:J'+y+')').style(Style_);
    worksheet.cell(startRow,11).formula('SUM(K10:K'+y+')').style(Style_);
    

    worksheet.cell(startRow,12).formula('COUNTIF(L10:L'+y+', "екзам")').style(Style_);
    worksheet.cell(startRow,13).style(Style_);
    worksheet.cell(startRow,14).style(Style_);
    worksheet.cell(startRow,15).style(Style_);
    worksheet.cell(startRow,16).style(Style_);
    worksheet.cell(startRow,17).style(Style_);
    startRow++;

    if (VK_disciplines.length === 0) {
      const D={
        alignment:{
          horizontal: 'center',
          vertical: 'center',
          wrapText: true,
        },
        font:{
          name:'TimeNewRoman'
        }
      };
      startRow++;
      worksheet.cell(startRow,1,startRow,3,true).string("Здобувач освіти__________").style(Style0);
      worksheet.cell(startRow,4,startRow,11,true).string("Відповідальна особа від закладу ФПО__________").style(Style0);
      worksheet.cell(startRow,13,startRow,16,true).string("Завідувач відділення_______________").style(Style0);
      worksheet.cell(startRow+1,3).string("(Підпис)").style(D);
      worksheet.cell(startRow+1,15,startRow+1,16,true).string("(Підпис)").style(D);
      worksheet.cell(startRow+1,17).string("(Підпис)").style(D);
    } else {
      worksheet.cell(startRow,1,startRow,17,true).string("Вибіркові освітні компоненти").style(Style);
      startRow++;
      VK_disciplines.forEach(discipline=>{
        Object.keys(discipline).forEach((key,index)=>{
          worksheet.cell(startRow,startColumn).number(C).style(Style_);
          worksheet.cell(startRow,startColumn+1).string(discipline.conVK).style(Style_);
          worksheet.cell(startRow,startColumn+2).string(discipline.name).style(Style_1);
          worksheet.cell(startRow,startColumn+3).number(discipline.countCredit).style(Style_);
          worksheet.cell(startRow,startColumn+4).number(discipline.hoursInWeek*discipline.countWeek).style(Style_);
          worksheet.cell(startRow,startColumn+5).number(discipline.lectures+discipline.practicalLaboratory+discipline.seminar).style(Style_);
          worksheet.cell(startRow,startColumn+6).number(discipline.lectures).style(Style_);
          worksheet.cell(startRow,startColumn+7).number(discipline.practicalLaboratory).style(Style_);
          worksheet.cell(startRow,startColumn+8).number(discipline.seminar).style(Style_);
          worksheet.cell(startRow,startColumn+9).number(discipline.countCredit*30-discipline.hoursInWeek*discipline.countWeek).style(Style_);
          worksheet.cell(startRow,startColumn+11).string(discipline.formOfControl).style(Style_);
          worksheet.cell(startRow,startColumn+10).style(Style_);
          worksheet.cell(startRow,startColumn+12).style(Style_);
          worksheet.cell(startRow,startColumn+13).style(Style_);
          worksheet.cell(startRow,startColumn+14).style(Style_);
          worksheet.cell(startRow,startColumn+15).style(Style_);
          worksheet.cell(startRow,startColumn+16).style(Style_);
        });
        startRow++;
        C++;
      });
      const D={
        alignment:{
          horizontal: 'center',
          vertical: 'center',
          wrapText: true,
        },
        font:{
          name:'TimeNewRoman'
        }
      };
      let t=y+1;
      console.log(y);
      const x=startRow-1;
      worksheet.cell(startRow,1).style(Style_);
      worksheet.cell(startRow,2).style(Style_);
      worksheet.cell(startRow,3).string("Всього:").style(Style_1);
      worksheet.cell(startRow,4).formula('SUM(D'+t+':D'+x+')').style(Style_);
      worksheet.cell(startRow,5).formula('SUM(E'+t+':E'+x+')').style(Style_);
      worksheet.cell(startRow,6).formula('SUM(F'+t+':F'+x+')').style(Style_);
      worksheet.cell(startRow,7).formula('SUM(G'+t+':G'+x+')').style(Style_);
      worksheet.cell(startRow,8).formula('SUM(H'+t+':H'+x+')').style(Style_);
      worksheet.cell(startRow,9).formula('SUM(I'+t+':I'+x+')').style(Style_);
      worksheet.cell(startRow,10).formula('SUM(J'+t+':J'+x+')').style(Style_);
      worksheet.cell(startRow,11).formula('SUM(K'+t+':K'+x+')').style(Style_);
      worksheet.cell(startRow,12).formula('COUNTIF(L10:L'+x+', "екзам")').style(Style_);
      worksheet.cell(startRow,13).style(Style_);
      worksheet.cell(startRow,14).style(Style_);
      worksheet.cell(startRow,15).style(Style_);
      worksheet.cell(startRow,16).style(Style_);
      worksheet.cell(startRow,17).style(Style_);


      startRow++;
      startRow++;
      worksheet.cell(startRow,1,startRow,3,true).string("Здобувач освіти__________").style(Style0);
      worksheet.cell(startRow,4,startRow,11,true).string("Відповідальна особа від закладу ФПО__________").style(Style0);
      worksheet.cell(startRow,13,startRow,16,true).string("Завідувач відділення_______________").style(Style0);
      worksheet.cell(startRow+1,3).string("(Підпис)").style(D);
      worksheet.cell(startRow+1,15,startRow+1,16,true).string("(Підпис)").style(D);
      worksheet.cell(startRow+1,17).string("(Підпис)").style(D);
    }

    two_semester(student,Ok_disciplines_2,VK_disciplines_2,workbook);
    three_semester(student,Ok_disciplines_3,VK_disciplines_3,workbook);
    four_semester(student,Ok_disciplines_4,VK_disciplines_4,workbook);
    five_semester(student,Ok_disciplines_5,VK_disciplines_5,workbook);
    six_semester(student,Ok_disciplines_6,VK_disciplines_6,workbook);
    Practice_stydent(workbook,Practices);
    const file_name=student[0].surename+" "+ student[0].name+" "+student[0].midle_name;
    // Зберігаємо книгу як файл Excel
    workbook.write(file_name+`.xlsx`);
  }

  function Title_page(workbook,student){
    var worksheet = workbook.addWorksheet('Титулка');
    const style_1={
      font:{
        size:14,
        name:"TimeNewRoman",
        bold:true,
      },
      alignment: {
        horizontal: 'center',
        vertical: 'center',
        wrapText: true,
      }
    }
    const style_2={
      font:{
        size:14,
        name:"TimeNewRoman",
      },
      alignment: {
        horizontal: 'center',
        wrapText: true,
      },
      border: { 
        left: {
          style: 'thin',
          color: 'black'
        },
        right: {
          style: 'thin',
          color: 'black'
        },
        top: {
          style: 'thin',
          color: 'black'
        },
        bottom: {
          style: 'thin',
          color: 'black'
        }
      }
    }
    const style_3={
      font:{
        size:12,
        name:"TimeNewRoman",
      },
      alignment: {
        horizontal: 'center',
        wrapText: true,
      },
      border: { 
        bottom: {
          style: 'thin',
          color: 'black'
        }
      }
    }
    const style_3_1={
      font:{
        size:12,
        name:"TimeNewRoman",
      },
      alignment: {
        horizontal: 'center',
        wrapText: true,
      },
    }
    const style_4={
      font:{
        size:13,
        name:"TimeNewRoman",
      },
    }
    const style_5={
      font:{
        size:12,
        name:"TimeNewRoman",
      },
      alignment: {
        horizontal: 'center',
        wrapText: true,
      },
    }
    worksheet.cell(1,1,1,6,true).string("ЧЕРВОНОГРАДСЬКИЙ ГІРНИЧО-ЕКОНОМІЧНИЙ ФАХОВИЙ КОЛЕДЖ").style(style_1);
    worksheet.cell(2,1,2,6,true).string("НАЗВА ВІДДІЛЕННЯ").style(style_1).style(style_1);
    worksheet.cell(3,1,3,6,true).string("ІНДИВІДУАЛЬНИЙ НАВЧАЛЬНИЙ ПЛАН").style(style_1).style(style_1);
    worksheet.cell(4,1,4,6,true).string("ЗДОБУВАЧА ФАХОВОЇ ПЕРЕДВИЩОЇ ОСВІТИ").style(style_1).style(style_1);
    worksheet.cell(5,1,5,6,true).string("на "+ student[0].year+"/"+(student[0].year+3)+" навчальний рік").style(style_1).style(style_1);
    worksheet.cell(7,1,10,1,true).string("3х4 \nМП").style(style_1).style(style_2);
    worksheet.column(1).setWidth(20);
    worksheet.row(1).setHeight(35);
    worksheet.row(7).setHeight(55);
    worksheet.cell(8,3,8,6,true).string(student[0].surename+" "+student[0].name+" "+student[0].midle_name).style(style_3);
    worksheet.cell(9,3,9,6,true).string("Прізвище, ім’я, по батькові здобувача освіти").style(style_3_1);
    worksheet.cell(12,1).string("Галузь знань").style(style_4);
    worksheet.cell(12,2,12,6,true).style(style_3);
    worksheet.cell(14,1).string("Спеціальність").style(style_4);
    worksheet.cell(14,2,14,6,true).string("Інженерія Програмного Забезпечення").style(style_3);
    worksheet.cell(16,1,16,2,true).string("Освітньо-професійна програма").style(style_4);
    worksheet.cell(16,3,16,6,true).style(style_3);
    worksheet.cell(18,1,18,2,true).string("Освітньо-професійний ступінь").style(style_4);
    worksheet.cell(18,3,18,6,true).style(style_3);
    worksheet.cell(20,1,20,2,true).string("Форма здобуття освіти").style(style_4);
    worksheet.cell(20,3,20,6,true).string("Денна форма").style(style_3);
    worksheet.cell(22,1,22,6,true).string("Зарахований(на) на II курс Наказ від «____» "+"____ "+student[0].year+"р."+" № ___ ").style(style_4);
    worksheet.cell(24,1,24,6,true).string("Договір про надання освітніх послуг від «____»"+"____ "+student[0].year+"р."+" № ___ ").style(style_4);
    worksheet.cell(26,1).string("Завідувач відділення").style(style_4);
    worksheet.cell(26,2,26,6,true).style(style_3);
    worksheet.cell(27,2,27,3,true).string("(Підпис)").style(style_5);
    worksheet.cell(27,4,27,6,true).string("(Ініціали, прізвище)").style(style_5);
    worksheet.cell(29,1,29,3,true).string("Відповідальна особа від закладу освіти").style(style_4);
    worksheet.cell(29,4,29,6,true).style(style_3);
    worksheet.cell(30,4).string("(Підпис)").style(style_5);
    worksheet.cell(30,5,30,6,true).string("(Ініціали, прізвище)").style(style_5);
    worksheet.cell(32,1,32,3,true).string("Здобувач фахової передвищої освіти").style(style_4);
    worksheet.cell(32,4,32,6,true).style(style_3);
    worksheet.cell(33,4).string("(Підпис)").style(style_5);
    worksheet.cell(33,5,33,6,true).string("(Ініціали, прізвище)").style(style_5);
  }

  function two_semester(student,Ok_disciplines,VK_disciplines,workbook){
    // Додаємо робочі аркуші до книги
    var worksheet = workbook.addWorksheet('II курс II семестер');

    const Style0={
      font: {
        name: 'Times New Roman',
        size:12
      },
    };
  worksheet.cell(1,1,1,3,true).string("Навчальний рік "+student[0].year+"/"+(student[0].year+1)).style(Style0);
  worksheet.cell(2,1,2,3,true).string("Курс: II").style(Style0);
  worksheet.cell(3,1,3,8,true).string("Семестр: II з_____"+(student[0].year+1)+"р. до_____"+(student[0].year+1)+"р.").style(Style0);
  worksheet.cell(4,1,4,8,true).string("Екзаменаційна сесія з_____"+(student[0].year+1)+"р. до_____"+(student[0].year+1)+"р.").style(Style0);
  worksheet.cell(1,8,1,17,true).string("Прізвище, ім’я по батькові здобувача освіти: "+student[0].surename+" "+student[0].name+" "+student[0].midle_name).style(Style0);
  worksheet.cell(2,8,2,17,true).string("Група:               1-ІПЗ-"+(student[0].year-1)%100).style(Style0);
  worksheet.column(1).setWidth(5);
  worksheet.column(2).setWidth(5);
  worksheet.column(4).setWidth(5);
  worksheet.column(5).setWidth(5);
  worksheet.column(6).setWidth(5);
  worksheet.column(7).setWidth(5);
  worksheet.column(8).setWidth(5);
  worksheet.column(9).setWidth(5);
  worksheet.column(10).setWidth(5);
  worksheet.column(11).setWidth(5);
  worksheet.column(12).setWidth(5);
  const Style={
    alignment: {
      horizontal: 'center',
      vertical: 'center',
      wrapText: true,
    },
    font: {
      name: 'Times New Roman',
      bold: true,
      size:10
    },
    border: { 
      left: {
        style: 'thin',
        color: 'black'
      },
      right: {
        style: 'thin',
        color: 'black'
      },
      top: {
        style: 'thin',
        color: 'black'
      },
      bottom: {
        style: 'thin',
        color: 'black'
      }
    }
  };
  const Style_rotate={
    alignment:{
      textRotation: 90,
      horizontal: 'center',
      vertical: 'center',
      wrapText: true,
    },
    font: {
      name: 'Times New Roman',
      bold: true,
      size:10
    },
    border: { 
      left: {
        style: 'thin',
        color: 'black'
      },
      right: {
        style: 'thin',
        color: 'black'
      },
      top: {
        style: 'thin',
        color: 'black'
      },
      bottom: {
        style: 'thin',
        color: 'black'
      }
    }
  };
    worksheet.cell(6,1,8,1,true).string("№ з/п").style(Style);
    worksheet.cell(6,2,8,2,true).string("Код ОК").style(Style);
    worksheet.cell(6,3,8,3,true).string("Назва\n освітнього\n компонента").style(Style);
    worksheet.column(3).setWidth(20);
    worksheet.cell(6,4,8,4,true).string("Кількість кредитів ЄКТС").style(Style_rotate);
    worksheet.row(6).setHeight(20);
    worksheet.row(7).setHeight(70);
    worksheet.row(8).setHeight(90);
    worksheet.cell(6,5,6,11,true).string("Кількість годин").style(Style);
    worksheet.cell(7,5,8,5,true).string("Всього").style(Style_rotate);
    worksheet.cell(7,6,8,6,true).string("Аудиторні години").style(Style_rotate);
    worksheet.cell(7,7,7,11,true).string("З них").style(Style);
    worksheet.cell(8,7).string("Лекції").style(Style_rotate);
    worksheet.cell(8,8).string("Практичні,\nЛабораторні").style(Style_rotate);
    worksheet.cell(8,9).string("Семінарські\n заняття").style(Style_rotate);
    worksheet.cell(8,10).string("Самостійна\n робота").style(Style_rotate);
    worksheet.cell(8,11).string("Індивідуальна\n робота").style(Style_rotate);
    worksheet.cell(6,12,8,12,true).string("Форма підсумкового контролю").style(Style_rotate);
    worksheet.cell(6,13,7,14,true).string("Оцінка").style(Style);
    worksheet.cell(8,13).string("за 12- бальною\n шкалою").style(Style);
    worksheet.cell(8,14).string("за шкалою\n закладу освіти").style(Style);
    worksheet.column(13).setWidth(8);
    worksheet.column(14).setWidth(8);
    worksheet.cell(6,15,8,15,true).string("Дата\n проведення\n \nпідсумкового\n контролю").style(Style);
    worksheet.column(15).setWidth(11);
    worksheet.cell(6,16,8,16,true).string("Прізвище,\n ініціали\n викладача").style(Style);
    worksheet.column(16).setWidth(9);
    worksheet.cell(6,17,8,17,true).string("Підпис\n викладача").style(Style);
    worksheet.cell(9,1,9,17,true).string("Обов’язкові освітні компоненти").style(Style);
    worksheet.column(17).setWidth(9);
    const Style_={
      alignment:{
        horizontal: 'center',
        vertical: 'center',
        wrapText: true,
      },
      font: {
        name: 'Times New Roman',
        size:10
      },
      border: { 
        left: {
          style: 'thin',
          color: 'black'
        },
        right: {
          style: 'thin',
          color: 'black'
        },
        top: {
          style: 'thin',
          color: 'black'
        },
        bottom: {
          style: 'thin',
          color: 'black'
        }
      }
    };
    const Style_1={
      alignment:{
        vertical: 'center',
        wrapText: true,
      },
      font: {
        name: 'Times New Roman',
        size:10
      },
      border: { 
        left: {
          style: 'thin',
          color: 'black'
        },
        right: {
          style: 'thin',
          color: 'black'
        },
        top: {
          style: 'thin',
          color: 'black'
        },
        bottom: {
          style: 'thin',
          color: 'black'
        }
      }
    };
    let startRow = 10;
    let C=1;
    let startColumn = 1;
    Ok_disciplines.forEach(discipline=>{
      Object.keys(discipline).forEach((key,index)=>{
        worksheet.cell(startRow,startColumn).number(C).style(Style_);
        worksheet.cell(startRow,startColumn+1).string(discipline.conVK).style(Style_);
        worksheet.cell(startRow,startColumn+2).string(discipline.name).style(Style_1);
        worksheet.cell(startRow,startColumn+3).number(discipline.countCredit).style(Style_);
        worksheet.cell(startRow,startColumn+4).number(discipline.hoursInWeek*discipline.countWeek).style(Style_);
        worksheet.cell(startRow,startColumn+5).number(discipline.lectures+discipline.practicalLaboratory+discipline.seminar).style(Style_);
        worksheet.cell(startRow,startColumn+6).number(discipline.lectures).style(Style_);
        worksheet.cell(startRow,startColumn+7).number(discipline.practicalLaboratory).style(Style_);
        worksheet.cell(startRow,startColumn+8).number(discipline.seminar).style(Style_);
        worksheet.cell(startRow,startColumn+9).number(discipline.countCredit*30-discipline.hoursInWeek*discipline.countWeek).style(Style_);
        worksheet.cell(startRow,startColumn+11).string(discipline.formOfControl).style(Style_);
        worksheet.cell(startRow,startColumn+10).style(Style_);
        worksheet.cell(startRow,startColumn+12).style(Style_);
        worksheet.cell(startRow,startColumn+13).style(Style_);
        worksheet.cell(startRow,startColumn+14).style(Style_);
        worksheet.cell(startRow,startColumn+15).style(Style_);
        worksheet.cell(startRow,startColumn+16).style(Style_);
      });
      startRow++;
      C++;
    });
    const y=startRow-1;
    worksheet.cell(startRow,1).style(Style_);
    worksheet.cell(startRow,2).style(Style_);
    worksheet.cell(startRow,3).string("Всього:").style(Style_1);
    worksheet.cell(startRow,4).formula('SUM(D10:D'+y+')').style(Style_);
    worksheet.cell(startRow,5).formula('SUM(E10:E'+y+')').style(Style_);
    worksheet.cell(startRow,6).formula('SUM(F10:F'+y+')').style(Style_);
    worksheet.cell(startRow,7).formula('SUM(G10:G'+y+')').style(Style_);
    worksheet.cell(startRow,8).formula('SUM(H10:H'+y+')').style(Style_);
    worksheet.cell(startRow,9).formula('SUM(I10:I'+y+')').style(Style_);
    worksheet.cell(startRow,10).formula('SUM(J10:J'+y+')').style(Style_);
    worksheet.cell(startRow,11).formula('SUM(K10:K'+y+')').style(Style_);
    

    worksheet.cell(startRow,12).formula('COUNTIF(L10:L'+y+', "екзам")').style(Style_);
    worksheet.cell(startRow,13).style(Style_);
    worksheet.cell(startRow,14).style(Style_);
    worksheet.cell(startRow,15).style(Style_);
    worksheet.cell(startRow,16).style(Style_);
    worksheet.cell(startRow,17).style(Style_);
    startRow++;

    if (VK_disciplines.length === 0) {
      const D={
        alignment:{
          horizontal: 'center',
          vertical: 'center',
          wrapText: true,
        },
        font:{
          name:'TimeNewRoman'
        }
      };
      startRow++;
      worksheet.cell(startRow,1,startRow,3,true).string("Здобувач освіти__________").style(Style0);
      worksheet.cell(startRow,4,startRow,11,true).string("Відповідальна особа від закладу ФПО__________").style(Style0);
      worksheet.cell(startRow,13,startRow,16,true).string("Завідувач відділення_______________").style(Style0);
      worksheet.cell(startRow+1,3).string("(Підпис)").style(D);
      worksheet.cell(startRow+1,15,startRow+1,16,true).string("(Підпис)").style(D);
      worksheet.cell(startRow+1,17).string("(Підпис)").style(D);
    } else {
      worksheet.cell(startRow,1,startRow,17,true).string("Вибіркові освітні компоненти").style(Style);
      startRow++;
      VK_disciplines.forEach(discipline=>{
        Object.keys(discipline).forEach((key,index)=>{
          worksheet.cell(startRow,startColumn).number(C).style(Style_);
          worksheet.cell(startRow,startColumn+1).string(discipline.conVK).style(Style_);
          worksheet.cell(startRow,startColumn+2).string(discipline.name).style(Style_1);
          worksheet.cell(startRow,startColumn+3).number(discipline.countCredit).style(Style_);
          worksheet.cell(startRow,startColumn+4).number(discipline.hoursInWeek*discipline.countWeek).style(Style_);
          worksheet.cell(startRow,startColumn+5).number(discipline.lectures+discipline.practicalLaboratory+discipline.seminar).style(Style_);
          worksheet.cell(startRow,startColumn+6).number(discipline.lectures).style(Style_);
          worksheet.cell(startRow,startColumn+7).number(discipline.practicalLaboratory).style(Style_);
          worksheet.cell(startRow,startColumn+8).number(discipline.seminar).style(Style_);
          worksheet.cell(startRow,startColumn+9).number(discipline.countCredit*30-discipline.hoursInWeek*discipline.countWeek).style(Style_);
          worksheet.cell(startRow,startColumn+11).string(discipline.formOfControl).style(Style_);
          worksheet.cell(startRow,startColumn+10).style(Style_);
          worksheet.cell(startRow,startColumn+12).style(Style_);
          worksheet.cell(startRow,startColumn+13).style(Style_);
          worksheet.cell(startRow,startColumn+14).style(Style_);
          worksheet.cell(startRow,startColumn+15).style(Style_);
          worksheet.cell(startRow,startColumn+16).style(Style_);
        });
        startRow++;
        C++;
      });
      const D={
        alignment:{
          horizontal: 'center',
          vertical: 'center',
          wrapText: true,
        },
        font:{
          name:'TimeNewRoman'
        }
      };
      let t=y+1;
      const x=startRow-1;
      worksheet.cell(startRow,1).style(Style_);
      worksheet.cell(startRow,2).style(Style_);
      worksheet.cell(startRow,3).string("Всього:").style(Style_1);
      worksheet.cell(startRow,4).formula('SUM(D'+t+':D'+x+')').style(Style_);
      worksheet.cell(startRow,5).formula('SUM(E'+t+':E'+x+')').style(Style_);
      worksheet.cell(startRow,6).formula('SUM(F'+t+':F'+x+')').style(Style_);
      worksheet.cell(startRow,7).formula('SUM(G'+t+':G'+x+')').style(Style_);
      worksheet.cell(startRow,8).formula('SUM(H'+t+':H'+x+')').style(Style_);
      worksheet.cell(startRow,9).formula('SUM(I'+t+':I'+x+')').style(Style_);
      worksheet.cell(startRow,10).formula('SUM(J'+t+':J'+x+')').style(Style_);
      worksheet.cell(startRow,11).formula('SUM(K'+t+':K'+x+')').style(Style_);
      worksheet.cell(startRow,12).formula('COUNTIF(L10:L'+x+', "екзам")').style(Style_);
      worksheet.cell(startRow,13).style(Style_);
      worksheet.cell(startRow,14).style(Style_);
      worksheet.cell(startRow,15).style(Style_);
      worksheet.cell(startRow,16).style(Style_);
      worksheet.cell(startRow,17).style(Style_);


      startRow++;
      startRow++;
      worksheet.cell(startRow,1,startRow,3,true).string("Здобувач освіти__________").style(Style0);
      worksheet.cell(startRow,4,startRow,11,true).string("Відповідальна особа від закладу ФПО__________").style(Style0);
      worksheet.cell(startRow,13,startRow,16,true).string("Завідувач відділення_______________").style(Style0);
      worksheet.cell(startRow+1,3).string("(Підпис)").style(D);
      worksheet.cell(startRow+1,15,startRow+1,16,true).string("(Підпис)").style(D);
      worksheet.cell(startRow+1,17).string("(Підпис)").style(D);
    }
  }

  function three_semester(student,Ok_disciplines,VK_disciplines,workbook){
    // Додаємо робочі аркуші до книги
    var worksheet = workbook.addWorksheet('III курс III семестер');

    const Style0={
      font: {
        name: 'Times New Roman',
        size:12
      },
    };
  worksheet.cell(1,1,1,3,true).string("Навчальний рік "+(student[0].year+1)+"/"+(student[0].year+2)).style(Style0);
  worksheet.cell(2,1,2,3,true).string("Курс: III").style(Style0);
  worksheet.cell(3,1,3,8,true).string("Семестр: III з_____"+(student[0].year+1)+"р. до_____"+(student[0].year+1)+"р.").style(Style0);
  worksheet.cell(4,1,4,8,true).string("Екзаменаційна сесія з_____"+(student[0].year+1)+"р. до_____"+(student[0].year+1)+"р.").style(Style0);
  worksheet.cell(1,8,1,17,true).string("Прізвище, ім’я по батькові здобувача освіти: "+student[0].surename+" "+student[0].name+" "+student[0].midle_name).style(Style0);
  worksheet.cell(2,8,2,17,true).string("Група:               1-ІПЗ-"+(student[0].year-1)%100).style(Style0);
  worksheet.column(1).setWidth(5);
  worksheet.column(2).setWidth(5);
  worksheet.column(4).setWidth(5);
  worksheet.column(5).setWidth(5);
  worksheet.column(6).setWidth(5);
  worksheet.column(7).setWidth(5);
  worksheet.column(8).setWidth(5);
  worksheet.column(9).setWidth(5);
  worksheet.column(10).setWidth(5);
  worksheet.column(11).setWidth(5);
  worksheet.column(12).setWidth(5);
  const Style={
    alignment: {
      horizontal: 'center',
      vertical: 'center',
      wrapText: true,
    },
    font: {
      name: 'Times New Roman',
      bold: true,
      size:10
    },
    border: { 
      left: {
        style: 'thin',
        color: 'black'
      },
      right: {
        style: 'thin',
        color: 'black'
      },
      top: {
        style: 'thin',
        color: 'black'
      },
      bottom: {
        style: 'thin',
        color: 'black'
      }
    }
  };
  const Style_rotate={
    alignment:{
      textRotation: 90,
      horizontal: 'center',
      vertical: 'center',
      wrapText: true,
    },
    font: {
      name: 'Times New Roman',
      bold: true,
      size:10
    },
    border: { 
      left: {
        style: 'thin',
        color: 'black'
      },
      right: {
        style: 'thin',
        color: 'black'
      },
      top: {
        style: 'thin',
        color: 'black'
      },
      bottom: {
        style: 'thin',
        color: 'black'
      }
    }
  };
    worksheet.cell(6,1,8,1,true).string("№ з/п").style(Style);
    worksheet.cell(6,2,8,2,true).string("Код ОК").style(Style);
    worksheet.cell(6,3,8,3,true).string("Назва\n освітнього\n компонента").style(Style);
    worksheet.column(3).setWidth(20);
    worksheet.cell(6,4,8,4,true).string("Кількість кредитів ЄКТС").style(Style_rotate);
    worksheet.row(6).setHeight(20);
    worksheet.row(7).setHeight(70);
    worksheet.row(8).setHeight(90);
    worksheet.cell(6,5,6,11,true).string("Кількість годин").style(Style);
    worksheet.cell(7,5,8,5,true).string("Всього").style(Style_rotate);
    worksheet.cell(7,6,8,6,true).string("Аудиторні години").style(Style_rotate);
    worksheet.cell(7,7,7,11,true).string("З них").style(Style);
    worksheet.cell(8,7).string("Лекції").style(Style_rotate);
    worksheet.cell(8,8).string("Практичні,\nЛабораторні").style(Style_rotate);
    worksheet.cell(8,9).string("Семінарські\n заняття").style(Style_rotate);
    worksheet.cell(8,10).string("Самостійна\n робота").style(Style_rotate);
    worksheet.cell(8,11).string("Індивідуальна\n робота").style(Style_rotate);
    worksheet.cell(6,12,8,12,true).string("Форма підсумкового контролю").style(Style_rotate);
    worksheet.cell(6,13,7,14,true).string("Оцінка").style(Style);
    worksheet.cell(8,13).string("за 12- бальною\n шкалою").style(Style);
    worksheet.cell(8,14).string("за шкалою\n закладу освіти").style(Style);
    worksheet.column(13).setWidth(8);
    worksheet.column(14).setWidth(8);
    worksheet.cell(6,15,8,15,true).string("Дата\n проведення\n \nпідсумкового\n контролю").style(Style);
    worksheet.column(15).setWidth(11);
    worksheet.cell(6,16,8,16,true).string("Прізвище,\n ініціали\n викладача").style(Style);
    worksheet.column(16).setWidth(9);
    worksheet.cell(6,17,8,17,true).string("Підпис\n викладача").style(Style);
    worksheet.cell(9,1,9,17,true).string("Обов’язкові освітні компоненти").style(Style);
    worksheet.column(17).setWidth(9);
    const Style_={
      alignment:{
        horizontal: 'center',
        vertical: 'center',
        wrapText: true,
      },
      font: {
        name: 'Times New Roman',
        size:10
      },
      border: { 
        left: {
          style: 'thin',
          color: 'black'
        },
        right: {
          style: 'thin',
          color: 'black'
        },
        top: {
          style: 'thin',
          color: 'black'
        },
        bottom: {
          style: 'thin',
          color: 'black'
        }
      }
    };
    const Style_1={
      alignment:{
        vertical: 'center',
        wrapText: true,
      },
      font: {
        name: 'Times New Roman',
        size:10
      },
      border: { 
        left: {
          style: 'thin',
          color: 'black'
        },
        right: {
          style: 'thin',
          color: 'black'
        },
        top: {
          style: 'thin',
          color: 'black'
        },
        bottom: {
          style: 'thin',
          color: 'black'
        }
      }
    };
    let startRow = 10;
    let C=1;
    let startColumn = 1;
    Ok_disciplines.forEach(discipline=>{
      Object.keys(discipline).forEach((key,index)=>{
        worksheet.cell(startRow,startColumn).number(C).style(Style_);
        worksheet.cell(startRow,startColumn+1).string(discipline.conVK).style(Style_);
        worksheet.cell(startRow,startColumn+2).string(discipline.name).style(Style_1);
        worksheet.cell(startRow,startColumn+3).number(discipline.countCredit).style(Style_);
        worksheet.cell(startRow,startColumn+4).number(discipline.hoursInWeek*discipline.countWeek).style(Style_);
        worksheet.cell(startRow,startColumn+5).number(discipline.lectures+discipline.practicalLaboratory+discipline.seminar).style(Style_);
        worksheet.cell(startRow,startColumn+6).number(discipline.lectures).style(Style_);
        worksheet.cell(startRow,startColumn+7).number(discipline.practicalLaboratory).style(Style_);
        worksheet.cell(startRow,startColumn+8).number(discipline.seminar).style(Style_);
        worksheet.cell(startRow,startColumn+9).number(discipline.countCredit*30-discipline.hoursInWeek*discipline.countWeek).style(Style_);
        worksheet.cell(startRow,startColumn+11).string(discipline.formOfControl).style(Style_);
        worksheet.cell(startRow,startColumn+10).style(Style_);
        worksheet.cell(startRow,startColumn+12).style(Style_);
        worksheet.cell(startRow,startColumn+13).style(Style_);
        worksheet.cell(startRow,startColumn+14).style(Style_);
        worksheet.cell(startRow,startColumn+15).style(Style_);
        worksheet.cell(startRow,startColumn+16).style(Style_);
      });
      startRow++;
      C++;
    });
    const y=startRow-1;
    worksheet.cell(startRow,1).style(Style_);
    worksheet.cell(startRow,2).style(Style_);
    worksheet.cell(startRow,3).string("Всього:").style(Style_1);
    worksheet.cell(startRow,4).formula('SUM(D10:D'+y+')').style(Style_);
    worksheet.cell(startRow,5).formula('SUM(E10:E'+y+')').style(Style_);
    worksheet.cell(startRow,6).formula('SUM(F10:F'+y+')').style(Style_);
    worksheet.cell(startRow,7).formula('SUM(G10:G'+y+')').style(Style_);
    worksheet.cell(startRow,8).formula('SUM(H10:H'+y+')').style(Style_);
    worksheet.cell(startRow,9).formula('SUM(I10:I'+y+')').style(Style_);
    worksheet.cell(startRow,10).formula('SUM(J10:J'+y+')').style(Style_);
    worksheet.cell(startRow,11).formula('SUM(K10:K'+y+')').style(Style_);
    

    worksheet.cell(startRow,12).formula('COUNTIF(L10:L'+y+', "екзам")').style(Style_);
    worksheet.cell(startRow,13).style(Style_);
    worksheet.cell(startRow,14).style(Style_);
    worksheet.cell(startRow,15).style(Style_);
    worksheet.cell(startRow,16).style(Style_);
    worksheet.cell(startRow,17).style(Style_);
    startRow++;

    if (VK_disciplines.length === 0) {
      const D={
        alignment:{
          horizontal: 'center',
          vertical: 'center',
          wrapText: true,
        },
        font:{
          name:'TimeNewRoman'
        }
      };
      startRow++;
      worksheet.cell(startRow,1,startRow,3,true).string("Здобувач освіти__________").style(Style0);
      worksheet.cell(startRow,4,startRow,11,true).string("Відповідальна особа від закладу ФПО__________").style(Style0);
      worksheet.cell(startRow,13,startRow,16,true).string("Завідувач відділення_______________").style(Style0);
      worksheet.cell(startRow+1,3).string("(Підпис)").style(D);
      worksheet.cell(startRow+1,15,startRow+1,16,true).string("(Підпис)").style(D);
      worksheet.cell(startRow+1,17).string("(Підпис)").style(D);
    } else {
      worksheet.cell(startRow,1,startRow,17,true).string("Вибіркові освітні компоненти").style(Style);
      startRow++;
      VK_disciplines.forEach(discipline=>{
        Object.keys(discipline).forEach((key,index)=>{
          worksheet.cell(startRow,startColumn).number(C).style(Style_);
          worksheet.cell(startRow,startColumn+1).string(discipline.conVK).style(Style_);
          worksheet.cell(startRow,startColumn+2).string(discipline.name).style(Style_1);
          worksheet.cell(startRow,startColumn+3).number(discipline.countCredit).style(Style_);
          worksheet.cell(startRow,startColumn+4).number(discipline.hoursInWeek*discipline.countWeek).style(Style_);
          worksheet.cell(startRow,startColumn+5).number(discipline.lectures+discipline.practicalLaboratory+discipline.seminar).style(Style_);
          worksheet.cell(startRow,startColumn+6).number(discipline.lectures).style(Style_);
          worksheet.cell(startRow,startColumn+7).number(discipline.practicalLaboratory).style(Style_);
          worksheet.cell(startRow,startColumn+8).number(discipline.seminar).style(Style_);
          worksheet.cell(startRow,startColumn+9).number(discipline.countCredit*30-discipline.hoursInWeek*discipline.countWeek).style(Style_);
          worksheet.cell(startRow,startColumn+11).string(discipline.formOfControl).style(Style_);
          worksheet.cell(startRow,startColumn+10).style(Style_);
          worksheet.cell(startRow,startColumn+12).style(Style_);
          worksheet.cell(startRow,startColumn+13).style(Style_);
          worksheet.cell(startRow,startColumn+14).style(Style_);
          worksheet.cell(startRow,startColumn+15).style(Style_);
          worksheet.cell(startRow,startColumn+16).style(Style_);
        });
        startRow++;
        C++;
      });
      const D={
        alignment:{
          horizontal: 'center',
          vertical: 'center',
          wrapText: true,
        },
        font:{
          name:'TimeNewRoman'
        }
      };
      let t=y+1;
      const x=startRow-1;
      worksheet.cell(startRow,1).style(Style_);
      worksheet.cell(startRow,2).style(Style_);
      worksheet.cell(startRow,3).string("Всього:").style(Style_1);
      worksheet.cell(startRow,4).formula('SUM(D'+t+':D'+x+')').style(Style_);
      worksheet.cell(startRow,5).formula('SUM(E'+t+':E'+x+')').style(Style_);
      worksheet.cell(startRow,6).formula('SUM(F'+t+':F'+x+')').style(Style_);
      worksheet.cell(startRow,7).formula('SUM(G'+t+':G'+x+')').style(Style_);
      worksheet.cell(startRow,8).formula('SUM(H'+t+':H'+x+')').style(Style_);
      worksheet.cell(startRow,9).formula('SUM(I'+t+':I'+x+')').style(Style_);
      worksheet.cell(startRow,10).formula('SUM(J'+t+':J'+x+')').style(Style_);
      worksheet.cell(startRow,11).formula('SUM(K'+t+':K'+x+')').style(Style_);
      worksheet.cell(startRow,12).formula('COUNTIF(L10:L'+x+', "екзам")').style(Style_);
      worksheet.cell(startRow,13).style(Style_);
      worksheet.cell(startRow,14).style(Style_);
      worksheet.cell(startRow,15).style(Style_);
      worksheet.cell(startRow,16).style(Style_);
      worksheet.cell(startRow,17).style(Style_);


      startRow++;
      startRow++;
      worksheet.cell(startRow,1,startRow,3,true).string("Здобувач освіти__________").style(Style0);
      worksheet.cell(startRow,4,startRow,11,true).string("Відповідальна особа від закладу ФПО__________").style(Style0);
      worksheet.cell(startRow,13,startRow,16,true).string("Завідувач відділення_______________").style(Style0);
      worksheet.cell(startRow+1,3).string("(Підпис)").style(D);
      worksheet.cell(startRow+1,15,startRow+1,16,true).string("(Підпис)").style(D);
      worksheet.cell(startRow+1,17).string("(Підпис)").style(D);
    }
  }

  function four_semester(student,Ok_disciplines,VK_disciplines,workbook){
    // Додаємо робочі аркуші до книги
    var worksheet = workbook.addWorksheet('III курс IV семестер');

    const Style0={
      font: {
        name: 'Times New Roman',
        size:12
      },
    };
  worksheet.cell(1,1,1,3,true).string("Навчальний рік "+(student[0].year+1)+"/"+(student[0].year+2)).style(Style0);
  worksheet.cell(2,1,2,3,true).string("Курс: III").style(Style0);
  worksheet.cell(3,1,3,8,true).string("Семестр: IV з_____"+(student[0].year+2)+"р. до_____"+(student[0].year+2)+"р.").style(Style0);
  worksheet.cell(4,1,4,8,true).string("Екзаменаційна сесія з_____"+(student[0].year+2)+"р. до_____"+(student[0].year+2)+"р.").style(Style0);
  worksheet.cell(1,8,1,17,true).string("Прізвище, ім’я по батькові здобувача освіти: "+student[0].surename+" "+student[0].name+" "+student[0].midle_name).style(Style0);
  worksheet.cell(2,8,2,17,true).string("Група:               1-ІПЗ-"+(student[0].year-1)%100).style(Style0);
  worksheet.column(1).setWidth(5);
  worksheet.column(2).setWidth(5);
  worksheet.column(4).setWidth(5);
  worksheet.column(5).setWidth(5);
  worksheet.column(6).setWidth(5);
  worksheet.column(7).setWidth(5);
  worksheet.column(8).setWidth(5);
  worksheet.column(9).setWidth(5);
  worksheet.column(10).setWidth(5);
  worksheet.column(11).setWidth(5);
  worksheet.column(12).setWidth(5);
  const Style={
    alignment: {
      horizontal: 'center',
      vertical: 'center',
      wrapText: true,
    },
    font: {
      name: 'Times New Roman',
      bold: true,
      size:10
    },
    border: { 
      left: {
        style: 'thin',
        color: 'black'
      },
      right: {
        style: 'thin',
        color: 'black'
      },
      top: {
        style: 'thin',
        color: 'black'
      },
      bottom: {
        style: 'thin',
        color: 'black'
      }
    }
  };
  const Style_rotate={
    alignment:{
      textRotation: 90,
      horizontal: 'center',
      vertical: 'center',
      wrapText: true,
    },
    font: {
      name: 'Times New Roman',
      bold: true,
      size:10
    },
    border: { 
      left: {
        style: 'thin',
        color: 'black'
      },
      right: {
        style: 'thin',
        color: 'black'
      },
      top: {
        style: 'thin',
        color: 'black'
      },
      bottom: {
        style: 'thin',
        color: 'black'
      }
    }
  };
    worksheet.cell(6,1,8,1,true).string("№ з/п").style(Style);
    worksheet.cell(6,2,8,2,true).string("Код ОК").style(Style);
    worksheet.cell(6,3,8,3,true).string("Назва\n освітнього\n компонента").style(Style);
    worksheet.column(3).setWidth(20);
    worksheet.cell(6,4,8,4,true).string("Кількість кредитів ЄКТС").style(Style_rotate);
    worksheet.row(6).setHeight(20);
    worksheet.row(7).setHeight(70);
    worksheet.row(8).setHeight(90);
    worksheet.cell(6,5,6,11,true).string("Кількість годин").style(Style);
    worksheet.cell(7,5,8,5,true).string("Всього").style(Style_rotate);
    worksheet.cell(7,6,8,6,true).string("Аудиторні години").style(Style_rotate);
    worksheet.cell(7,7,7,11,true).string("З них").style(Style);
    worksheet.cell(8,7).string("Лекції").style(Style_rotate);
    worksheet.cell(8,8).string("Практичні,\nЛабораторні").style(Style_rotate);
    worksheet.cell(8,9).string("Семінарські\n заняття").style(Style_rotate);
    worksheet.cell(8,10).string("Самостійна\n робота").style(Style_rotate);
    worksheet.cell(8,11).string("Індивідуальна\n робота").style(Style_rotate);
    worksheet.cell(6,12,8,12,true).string("Форма підсумкового контролю").style(Style_rotate);
    worksheet.cell(6,13,7,14,true).string("Оцінка").style(Style);
    worksheet.cell(8,13).string("за 12- бальною\n шкалою").style(Style);
    worksheet.cell(8,14).string("за шкалою\n закладу освіти").style(Style);
    worksheet.column(13).setWidth(8);
    worksheet.column(14).setWidth(8);
    worksheet.cell(6,15,8,15,true).string("Дата\n проведення\n \nпідсумкового\n контролю").style(Style);
    worksheet.column(15).setWidth(11);
    worksheet.cell(6,16,8,16,true).string("Прізвище,\n ініціали\n викладача").style(Style);
    worksheet.column(16).setWidth(9);
    worksheet.cell(6,17,8,17,true).string("Підпис\n викладача").style(Style);
    worksheet.cell(9,1,9,17,true).string("Обов’язкові освітні компоненти").style(Style);
    worksheet.column(17).setWidth(9);
    const Style_={
      alignment:{
        horizontal: 'center',
        vertical: 'center',
        wrapText: true,
      },
      font: {
        name: 'Times New Roman',
        size:10
      },
      border: { 
        left: {
          style: 'thin',
          color: 'black'
        },
        right: {
          style: 'thin',
          color: 'black'
        },
        top: {
          style: 'thin',
          color: 'black'
        },
        bottom: {
          style: 'thin',
          color: 'black'
        }
      }
    };
    const Style_1={
      alignment:{
        vertical: 'center',
        wrapText: true,
      },
      font: {
        name: 'Times New Roman',
        size:10
      },
      border: { 
        left: {
          style: 'thin',
          color: 'black'
        },
        right: {
          style: 'thin',
          color: 'black'
        },
        top: {
          style: 'thin',
          color: 'black'
        },
        bottom: {
          style: 'thin',
          color: 'black'
        }
      }
    };
    let startRow = 10;
    let C=1;
    let startColumn = 1;
    Ok_disciplines.forEach(discipline=>{
      Object.keys(discipline).forEach((key,index)=>{
        worksheet.cell(startRow,startColumn).number(C).style(Style_);
        worksheet.cell(startRow,startColumn+1).string(discipline.conVK).style(Style_);
        worksheet.cell(startRow,startColumn+2).string(discipline.name).style(Style_1);
        worksheet.cell(startRow,startColumn+3).number(discipline.countCredit).style(Style_);
        worksheet.cell(startRow,startColumn+4).number(discipline.hoursInWeek*discipline.countWeek).style(Style_);
        worksheet.cell(startRow,startColumn+5).number(discipline.lectures+discipline.practicalLaboratory+discipline.seminar).style(Style_);
        worksheet.cell(startRow,startColumn+6).number(discipline.lectures).style(Style_);
        worksheet.cell(startRow,startColumn+7).number(discipline.practicalLaboratory).style(Style_);
        worksheet.cell(startRow,startColumn+8).number(discipline.seminar).style(Style_);
        worksheet.cell(startRow,startColumn+9).number(discipline.countCredit*30-discipline.hoursInWeek*discipline.countWeek).style(Style_);
        worksheet.cell(startRow,startColumn+11).string(discipline.formOfControl).style(Style_);
        worksheet.cell(startRow,startColumn+10).style(Style_);
        worksheet.cell(startRow,startColumn+12).style(Style_);
        worksheet.cell(startRow,startColumn+13).style(Style_);
        worksheet.cell(startRow,startColumn+14).style(Style_);
        worksheet.cell(startRow,startColumn+15).style(Style_);
        worksheet.cell(startRow,startColumn+16).style(Style_);
      });
      startRow++;
      C++;
    });
    const y=startRow-1;
    worksheet.cell(startRow,1).style(Style_);
    worksheet.cell(startRow,2).style(Style_);
    worksheet.cell(startRow,3).string("Всього:").style(Style_1);
    worksheet.cell(startRow,4).formula('SUM(D10:D'+y+')').style(Style_);
    worksheet.cell(startRow,5).formula('SUM(E10:E'+y+')').style(Style_);
    worksheet.cell(startRow,6).formula('SUM(F10:F'+y+')').style(Style_);
    worksheet.cell(startRow,7).formula('SUM(G10:G'+y+')').style(Style_);
    worksheet.cell(startRow,8).formula('SUM(H10:H'+y+')').style(Style_);
    worksheet.cell(startRow,9).formula('SUM(I10:I'+y+')').style(Style_);
    worksheet.cell(startRow,10).formula('SUM(J10:J'+y+')').style(Style_);
    worksheet.cell(startRow,11).formula('SUM(K10:K'+y+')').style(Style_);
    

    worksheet.cell(startRow,12).formula('COUNTIF(L10:L'+y+', "екзам")').style(Style_);
    worksheet.cell(startRow,13).style(Style_);
    worksheet.cell(startRow,14).style(Style_);
    worksheet.cell(startRow,15).style(Style_);
    worksheet.cell(startRow,16).style(Style_);
    worksheet.cell(startRow,17).style(Style_);
    startRow++;

    if (VK_disciplines.length === 0) {
      const D={
        alignment:{
          horizontal: 'center',
          vertical: 'center',
          wrapText: true,
        },
        font:{
          name:'TimeNewRoman'
        }
      };
      startRow++;
      worksheet.cell(startRow,1,startRow,3,true).string("Здобувач освіти__________").style(Style0);
      worksheet.cell(startRow,4,startRow,11,true).string("Відповідальна особа від закладу ФПО__________").style(Style0);
      worksheet.cell(startRow,13,startRow,16,true).string("Завідувач відділення_______________").style(Style0);
      worksheet.cell(startRow+1,3).string("(Підпис)").style(D);
      worksheet.cell(startRow+1,15,startRow+1,16,true).string("(Підпис)").style(D);
      worksheet.cell(startRow+1,17).string("(Підпис)").style(D);
    } else {
      worksheet.cell(startRow,1,startRow,17,true).string("Вибіркові освітні компоненти").style(Style);
      startRow++;
      VK_disciplines.forEach(discipline=>{
        Object.keys(discipline).forEach((key,index)=>{
          worksheet.cell(startRow,startColumn).number(C).style(Style_);
          worksheet.cell(startRow,startColumn+1).string(discipline.conVK).style(Style_);
          worksheet.cell(startRow,startColumn+2).string(discipline.name).style(Style_1);
          worksheet.cell(startRow,startColumn+3).number(discipline.countCredit).style(Style_);
          worksheet.cell(startRow,startColumn+4).number(discipline.hoursInWeek*discipline.countWeek).style(Style_);
          worksheet.cell(startRow,startColumn+5).number(discipline.lectures+discipline.practicalLaboratory+discipline.seminar).style(Style_);
          worksheet.cell(startRow,startColumn+6).number(discipline.lectures).style(Style_);
          worksheet.cell(startRow,startColumn+7).number(discipline.practicalLaboratory).style(Style_);
          worksheet.cell(startRow,startColumn+8).number(discipline.seminar).style(Style_);
          worksheet.cell(startRow,startColumn+9).number(discipline.countCredit*30-discipline.hoursInWeek*discipline.countWeek).style(Style_);
          worksheet.cell(startRow,startColumn+11).string(discipline.formOfControl).style(Style_);
          worksheet.cell(startRow,startColumn+10).style(Style_);
          worksheet.cell(startRow,startColumn+12).style(Style_);
          worksheet.cell(startRow,startColumn+13).style(Style_);
          worksheet.cell(startRow,startColumn+14).style(Style_);
          worksheet.cell(startRow,startColumn+15).style(Style_);
          worksheet.cell(startRow,startColumn+16).style(Style_);
        });
        startRow++;
        C++;
      });
      const D={
        alignment:{
          horizontal: 'center',
          vertical: 'center',
          wrapText: true,
        },
        font:{
          name:'TimeNewRoman'
        }
      };
      let t=y+1;
      const x=startRow-1;
      worksheet.cell(startRow,1).style(Style_);
      worksheet.cell(startRow,2).style(Style_);
      worksheet.cell(startRow,3).string("Всього:").style(Style_1);
      worksheet.cell(startRow,4).formula('SUM(D'+t+':D'+x+')').style(Style_);
      worksheet.cell(startRow,5).formula('SUM(E'+t+':E'+x+')').style(Style_);
      worksheet.cell(startRow,6).formula('SUM(F'+t+':F'+x+')').style(Style_);
      worksheet.cell(startRow,7).formula('SUM(G'+t+':G'+x+')').style(Style_);
      worksheet.cell(startRow,8).formula('SUM(H'+t+':H'+x+')').style(Style_);
      worksheet.cell(startRow,9).formula('SUM(I'+t+':I'+x+')').style(Style_);
      worksheet.cell(startRow,10).formula('SUM(J'+t+':J'+x+')').style(Style_);
      worksheet.cell(startRow,11).formula('SUM(K'+t+':K'+x+')').style(Style_);
      worksheet.cell(startRow,12).formula('COUNTIF(L10:L'+x+', "екзам")').style(Style_);
      worksheet.cell(startRow,13).style(Style_);
      worksheet.cell(startRow,14).style(Style_);
      worksheet.cell(startRow,15).style(Style_);
      worksheet.cell(startRow,16).style(Style_);
      worksheet.cell(startRow,17).style(Style_);


      startRow++;
      startRow++;
      worksheet.cell(startRow,1,startRow,3,true).string("Здобувач освіти__________").style(Style0);
      worksheet.cell(startRow,4,startRow,11,true).string("Відповідальна особа від закладу ФПО__________").style(Style0);
      worksheet.cell(startRow,13,startRow,16,true).string("Завідувач відділення_______________").style(Style0);
      worksheet.cell(startRow+1,3).string("(Підпис)").style(D);
      worksheet.cell(startRow+1,15,startRow+1,16,true).string("(Підпис)").style(D);
      worksheet.cell(startRow+1,17).string("(Підпис)").style(D);
    }
  }

  function five_semester(student,Ok_disciplines,VK_disciplines,workbook){
    // Додаємо робочі аркуші до книги
    var worksheet = workbook.addWorksheet('IV курс V семестер');

    const Style0={
      font: {
        name: 'Times New Roman',
        size:12
      },
    };
  worksheet.cell(1,1,1,3,true).string("Навчальний рік "+(student[0].year+2)+"/"+(student[0].year+3)).style(Style0);
  worksheet.cell(2,1,2,3,true).string("Курс: IV").style(Style0);
  worksheet.cell(3,1,3,8,true).string("Семестр: V з_____"+(student[0].year+2)+"р. до_____"+(student[0].year+2)+"р.").style(Style0);
  worksheet.cell(4,1,4,8,true).string("Екзаменаційна сесія з_____"+(student[0].year+2)+"р. до_____"+(student[0].year+2)+"р.").style(Style0);
  worksheet.cell(1,8,1,17,true).string("Прізвище, ім’я по батькові здобувача освіти: "+student[0].surename+" "+student[0].name+" "+student[0].midle_name).style(Style0);
  worksheet.cell(2,8,2,17,true).string("Група:               1-ІПЗ-"+(student[0].year-1)%100).style(Style0);
  worksheet.column(1).setWidth(5);
  worksheet.column(2).setWidth(5);
  worksheet.column(4).setWidth(5);
  worksheet.column(5).setWidth(5);
  worksheet.column(6).setWidth(5);
  worksheet.column(7).setWidth(5);
  worksheet.column(8).setWidth(5);
  worksheet.column(9).setWidth(5);
  worksheet.column(10).setWidth(5);
  worksheet.column(11).setWidth(5);
  worksheet.column(12).setWidth(5);
  const Style={
    alignment: {
      horizontal: 'center',
      vertical: 'center',
      wrapText: true,
    },
    font: {
      name: 'Times New Roman',
      bold: true,
      size:10
    },
    border: { 
      left: {
        style: 'thin',
        color: 'black'
      },
      right: {
        style: 'thin',
        color: 'black'
      },
      top: {
        style: 'thin',
        color: 'black'
      },
      bottom: {
        style: 'thin',
        color: 'black'
      }
    }
  };
  const Style_rotate={
    alignment:{
      textRotation: 90,
      horizontal: 'center',
      vertical: 'center',
      wrapText: true,
    },
    font: {
      name: 'Times New Roman',
      bold: true,
      size:10
    },
    border: { 
      left: {
        style: 'thin',
        color: 'black'
      },
      right: {
        style: 'thin',
        color: 'black'
      },
      top: {
        style: 'thin',
        color: 'black'
      },
      bottom: {
        style: 'thin',
        color: 'black'
      }
    }
  };
    worksheet.cell(6,1,8,1,true).string("№ з/п").style(Style);
    worksheet.cell(6,2,8,2,true).string("Код ОК").style(Style);
    worksheet.cell(6,3,8,3,true).string("Назва\n освітнього\n компонента").style(Style);
    worksheet.column(3).setWidth(20);
    worksheet.cell(6,4,8,4,true).string("Кількість кредитів ЄКТС").style(Style_rotate);
    worksheet.row(6).setHeight(20);
    worksheet.row(7).setHeight(70);
    worksheet.row(8).setHeight(90);
    worksheet.cell(6,5,6,11,true).string("Кількість годин").style(Style);
    worksheet.cell(7,5,8,5,true).string("Всього").style(Style_rotate);
    worksheet.cell(7,6,8,6,true).string("Аудиторні години").style(Style_rotate);
    worksheet.cell(7,7,7,11,true).string("З них").style(Style);
    worksheet.cell(8,7).string("Лекції").style(Style_rotate);
    worksheet.cell(8,8).string("Практичні,\nЛабораторні").style(Style_rotate);
    worksheet.cell(8,9).string("Семінарські\n заняття").style(Style_rotate);
    worksheet.cell(8,10).string("Самостійна\n робота").style(Style_rotate);
    worksheet.cell(8,11).string("Індивідуальна\n робота").style(Style_rotate);
    worksheet.cell(6,12,8,12,true).string("Форма підсумкового контролю").style(Style_rotate);
    worksheet.cell(6,13,7,14,true).string("Оцінка").style(Style);
    worksheet.cell(8,13).string("за 12- бальною\n шкалою").style(Style);
    worksheet.cell(8,14).string("за шкалою\n закладу освіти").style(Style);
    worksheet.column(13).setWidth(8);
    worksheet.column(14).setWidth(8);
    worksheet.cell(6,15,8,15,true).string("Дата\n проведення\n \nпідсумкового\n контролю").style(Style);
    worksheet.column(15).setWidth(11);
    worksheet.cell(6,16,8,16,true).string("Прізвище,\n ініціали\n викладача").style(Style);
    worksheet.column(16).setWidth(9);
    worksheet.cell(6,17,8,17,true).string("Підпис\n викладача").style(Style);
    worksheet.cell(9,1,9,17,true).string("Обов’язкові освітні компоненти").style(Style);
    worksheet.column(17).setWidth(9);
    const Style_={
      alignment:{
        horizontal: 'center',
        vertical: 'center',
        wrapText: true,
      },
      font: {
        name: 'Times New Roman',
        size:10
      },
      border: { 
        left: {
          style: 'thin',
          color: 'black'
        },
        right: {
          style: 'thin',
          color: 'black'
        },
        top: {
          style: 'thin',
          color: 'black'
        },
        bottom: {
          style: 'thin',
          color: 'black'
        }
      }
    };
    const Style_1={
      alignment:{
        vertical: 'center',
        wrapText: true,
      },
      font: {
        name: 'Times New Roman',
        size:10
      },
      border: { 
        left: {
          style: 'thin',
          color: 'black'
        },
        right: {
          style: 'thin',
          color: 'black'
        },
        top: {
          style: 'thin',
          color: 'black'
        },
        bottom: {
          style: 'thin',
          color: 'black'
        }
      }
    };
    let startRow = 10;
    let C=1;
    let startColumn = 1;
    Ok_disciplines.forEach(discipline=>{
      Object.keys(discipline).forEach((key,index)=>{
        worksheet.cell(startRow,startColumn).number(C).style(Style_);
        worksheet.cell(startRow,startColumn+1).string(discipline.conVK).style(Style_);
        worksheet.cell(startRow,startColumn+2).string(discipline.name).style(Style_1);
        worksheet.cell(startRow,startColumn+3).number(discipline.countCredit).style(Style_);
        worksheet.cell(startRow,startColumn+4).number(discipline.hoursInWeek*discipline.countWeek).style(Style_);
        worksheet.cell(startRow,startColumn+5).number(discipline.lectures+discipline.practicalLaboratory+discipline.seminar).style(Style_);
        worksheet.cell(startRow,startColumn+6).number(discipline.lectures).style(Style_);
        worksheet.cell(startRow,startColumn+7).number(discipline.practicalLaboratory).style(Style_);
        worksheet.cell(startRow,startColumn+8).number(discipline.seminar).style(Style_);
        worksheet.cell(startRow,startColumn+9).number(discipline.countCredit*30-discipline.hoursInWeek*discipline.countWeek).style(Style_);
        worksheet.cell(startRow,startColumn+11).string(discipline.formOfControl).style(Style_);
        worksheet.cell(startRow,startColumn+10).style(Style_);
        worksheet.cell(startRow,startColumn+12).style(Style_);
        worksheet.cell(startRow,startColumn+13).style(Style_);
        worksheet.cell(startRow,startColumn+14).style(Style_);
        worksheet.cell(startRow,startColumn+15).style(Style_);
        worksheet.cell(startRow,startColumn+16).style(Style_);
      });
      startRow++;
      C++;
    });
    const y=startRow-1;
    worksheet.cell(startRow,1).style(Style_);
    worksheet.cell(startRow,2).style(Style_);
    worksheet.cell(startRow,3).string("Всього:").style(Style_1);
    worksheet.cell(startRow,4).formula('SUM(D10:D'+y+')').style(Style_);
    worksheet.cell(startRow,5).formula('SUM(E10:E'+y+')').style(Style_);
    worksheet.cell(startRow,6).formula('SUM(F10:F'+y+')').style(Style_);
    worksheet.cell(startRow,7).formula('SUM(G10:G'+y+')').style(Style_);
    worksheet.cell(startRow,8).formula('SUM(H10:H'+y+')').style(Style_);
    worksheet.cell(startRow,9).formula('SUM(I10:I'+y+')').style(Style_);
    worksheet.cell(startRow,10).formula('SUM(J10:J'+y+')').style(Style_);
    worksheet.cell(startRow,11).formula('SUM(K10:K'+y+')').style(Style_);
    

    worksheet.cell(startRow,12).formula('COUNTIF(L10:L'+y+', "екзам")').style(Style_);
    worksheet.cell(startRow,13).style(Style_);
    worksheet.cell(startRow,14).style(Style_);
    worksheet.cell(startRow,15).style(Style_);
    worksheet.cell(startRow,16).style(Style_);
    worksheet.cell(startRow,17).style(Style_);
    startRow++;

    if (VK_disciplines.length === 0) {
      const D={
        alignment:{
          horizontal: 'center',
          vertical: 'center',
          wrapText: true,
        },
        font:{
          name:'TimeNewRoman'
        }
      };
      startRow++;
      worksheet.cell(startRow,1,startRow,3,true).string("Здобувач освіти__________").style(Style0);
      worksheet.cell(startRow,4,startRow,11,true).string("Відповідальна особа від закладу ФПО__________").style(Style0);
      worksheet.cell(startRow,13,startRow,16,true).string("Завідувач відділення_______________").style(Style0);
      worksheet.cell(startRow+1,3).string("(Підпис)").style(D);
      worksheet.cell(startRow+1,15,startRow+1,16,true).string("(Підпис)").style(D);
      worksheet.cell(startRow+1,17).string("(Підпис)").style(D);
    } else {
      worksheet.cell(startRow,1,startRow,17,true).string("Вибіркові освітні компоненти").style(Style);
      startRow++;
      VK_disciplines.forEach(discipline=>{
        Object.keys(discipline).forEach((key,index)=>{
          worksheet.cell(startRow,startColumn).number(C).style(Style_);
          worksheet.cell(startRow,startColumn+1).string(discipline.conVK).style(Style_);
          worksheet.cell(startRow,startColumn+2).string(discipline.name).style(Style_1);
          worksheet.cell(startRow,startColumn+3).number(discipline.countCredit).style(Style_);
          worksheet.cell(startRow,startColumn+4).number(discipline.hoursInWeek*discipline.countWeek).style(Style_);
          worksheet.cell(startRow,startColumn+5).number(discipline.lectures+discipline.practicalLaboratory+discipline.seminar).style(Style_);
          worksheet.cell(startRow,startColumn+6).number(discipline.lectures).style(Style_);
          worksheet.cell(startRow,startColumn+7).number(discipline.practicalLaboratory).style(Style_);
          worksheet.cell(startRow,startColumn+8).number(discipline.seminar).style(Style_);
          worksheet.cell(startRow,startColumn+9).number(discipline.countCredit*30-discipline.hoursInWeek*discipline.countWeek).style(Style_);
          worksheet.cell(startRow,startColumn+11).string(discipline.formOfControl).style(Style_);
          worksheet.cell(startRow,startColumn+10).style(Style_);
          worksheet.cell(startRow,startColumn+12).style(Style_);
          worksheet.cell(startRow,startColumn+13).style(Style_);
          worksheet.cell(startRow,startColumn+14).style(Style_);
          worksheet.cell(startRow,startColumn+15).style(Style_);
          worksheet.cell(startRow,startColumn+16).style(Style_);
        });
        startRow++;
        C++;
      });
      const D={
        alignment:{
          horizontal: 'center',
          vertical: 'center',
          wrapText: true,
        },
        font:{
          name:'TimeNewRoman'
        }
      };
      let t=y+1;
      const x=startRow-1;
      worksheet.cell(startRow,1).style(Style_);
      worksheet.cell(startRow,2).style(Style_);
      worksheet.cell(startRow,3).string("Всього:").style(Style_1);
      worksheet.cell(startRow,4).formula('SUM(D'+t+':D'+x+')').style(Style_);
      worksheet.cell(startRow,5).formula('SUM(E'+t+':E'+x+')').style(Style_);
      worksheet.cell(startRow,6).formula('SUM(F'+t+':F'+x+')').style(Style_);
      worksheet.cell(startRow,7).formula('SUM(G'+t+':G'+x+')').style(Style_);
      worksheet.cell(startRow,8).formula('SUM(H'+t+':H'+x+')').style(Style_);
      worksheet.cell(startRow,9).formula('SUM(I'+t+':I'+x+')').style(Style_);
      worksheet.cell(startRow,10).formula('SUM(J'+t+':J'+x+')').style(Style_);
      worksheet.cell(startRow,11).formula('SUM(K'+t+':K'+x+')').style(Style_);
      worksheet.cell(startRow,12).formula('COUNTIF(L10:L'+x+', "екзам")').style(Style_);
      worksheet.cell(startRow,13).style(Style_);
      worksheet.cell(startRow,14).style(Style_);
      worksheet.cell(startRow,15).style(Style_);
      worksheet.cell(startRow,16).style(Style_);
      worksheet.cell(startRow,17).style(Style_);


      startRow++;
      startRow++;
      worksheet.cell(startRow,1,startRow,3,true).string("Здобувач освіти__________").style(Style0);
      worksheet.cell(startRow,4,startRow,11,true).string("Відповідальна особа від закладу ФПО__________").style(Style0);
      worksheet.cell(startRow,13,startRow,16,true).string("Завідувач відділення_______________").style(Style0);
      worksheet.cell(startRow+1,3).string("(Підпис)").style(D);
      worksheet.cell(startRow+1,15,startRow+1,16,true).string("(Підпис)").style(D);
      worksheet.cell(startRow+1,17).string("(Підпис)").style(D);
    }
  }

  function six_semester(student,Ok_disciplines,VK_disciplines,workbook){
    // Додаємо робочі аркуші до книги
    var worksheet = workbook.addWorksheet('IV курс VI семестер');

    const Style0={
      font: {
        name: 'Times New Roman',
        size:12
      },
    };
  worksheet.cell(1,1,1,3,true).string("Навчальний рік "+(student[0].year+2)+"/"+(student[0].year+3)).style(Style0);
  worksheet.cell(2,1,2,3,true).string("Курс: IV").style(Style0);
  worksheet.cell(3,1,3,8,true).string("Семестр: VI з_____"+(student[0].year+3)+"р. до_____"+(student[0].year+3)+"р.").style(Style0);
  worksheet.cell(4,1,4,8,true).string("Екзаменаційна сесія з_____"+(student[0].year+3)+"р. до_____"+(student[0].year+3)+"р.").style(Style0);
  worksheet.cell(1,8,1,17,true).string("Прізвище, ім’я по батькові здобувача освіти: "+student[0].surename+" "+student[0].name+" "+student[0].midle_name).style(Style0);
  worksheet.cell(2,8,2,17,true).string("Група:               1-ІПЗ-"+(student[0].year-1)%100).style(Style0);
  worksheet.column(1).setWidth(5);
  worksheet.column(2).setWidth(5);
  worksheet.column(4).setWidth(5);
  worksheet.column(5).setWidth(5);
  worksheet.column(6).setWidth(5);
  worksheet.column(7).setWidth(5);
  worksheet.column(8).setWidth(5);
  worksheet.column(9).setWidth(5);
  worksheet.column(10).setWidth(5);
  worksheet.column(11).setWidth(5);
  worksheet.column(12).setWidth(5);
  const Style={
    alignment: {
      horizontal: 'center',
      vertical: 'center',
      wrapText: true,
    },
    font: {
      name: 'Times New Roman',
      bold: true,
      size:10
    },
    border: { 
      left: {
        style: 'thin',
        color: 'black'
      },
      right: {
        style: 'thin',
        color: 'black'
      },
      top: {
        style: 'thin',
        color: 'black'
      },
      bottom: {
        style: 'thin',
        color: 'black'
      }
    }
  };
  const Style_rotate={
    alignment:{
      textRotation: 90,
      horizontal: 'center',
      vertical: 'center',
      wrapText: true,
    },
    font: {
      name: 'Times New Roman',
      bold: true,
      size:10
    },
    border: { 
      left: {
        style: 'thin',
        color: 'black'
      },
      right: {
        style: 'thin',
        color: 'black'
      },
      top: {
        style: 'thin',
        color: 'black'
      },
      bottom: {
        style: 'thin',
        color: 'black'
      }
    }
  };
    worksheet.cell(6,1,8,1,true).string("№ з/п").style(Style);
    worksheet.cell(6,2,8,2,true).string("Код ОК").style(Style);
    worksheet.cell(6,3,8,3,true).string("Назва\n освітнього\n компонента").style(Style);
    worksheet.column(3).setWidth(20);
    worksheet.cell(6,4,8,4,true).string("Кількість кредитів ЄКТС").style(Style_rotate);
    worksheet.row(6).setHeight(20);
    worksheet.row(7).setHeight(70);
    worksheet.row(8).setHeight(90);
    worksheet.cell(6,5,6,11,true).string("Кількість годин").style(Style);
    worksheet.cell(7,5,8,5,true).string("Всього").style(Style_rotate);
    worksheet.cell(7,6,8,6,true).string("Аудиторні години").style(Style_rotate);
    worksheet.cell(7,7,7,11,true).string("З них").style(Style);
    worksheet.cell(8,7).string("Лекції").style(Style_rotate);
    worksheet.cell(8,8).string("Практичні,\nЛабораторні").style(Style_rotate);
    worksheet.cell(8,9).string("Семінарські\n заняття").style(Style_rotate);
    worksheet.cell(8,10).string("Самостійна\n робота").style(Style_rotate);
    worksheet.cell(8,11).string("Індивідуальна\n робота").style(Style_rotate);
    worksheet.cell(6,12,8,12,true).string("Форма підсумкового контролю").style(Style_rotate);
    worksheet.cell(6,13,7,14,true).string("Оцінка").style(Style);
    worksheet.cell(8,13).string("за 12- бальною\n шкалою").style(Style);
    worksheet.cell(8,14).string("за шкалою\n закладу освіти").style(Style);
    worksheet.column(13).setWidth(8);
    worksheet.column(14).setWidth(8);
    worksheet.cell(6,15,8,15,true).string("Дата\n проведення\n \nпідсумкового\n контролю").style(Style);
    worksheet.column(15).setWidth(11);
    worksheet.cell(6,16,8,16,true).string("Прізвище,\n ініціали\n викладача").style(Style);
    worksheet.column(16).setWidth(9);
    worksheet.cell(6,17,8,17,true).string("Підпис\n викладача").style(Style);
    worksheet.cell(9,1,9,17,true).string("Обов’язкові освітні компоненти").style(Style);
    worksheet.column(17).setWidth(9);
    const Style_={
      alignment:{
        horizontal: 'center',
        vertical: 'center',
        wrapText: true,
      },
      font: {
        name: 'Times New Roman',
        size:10
      },
      border: { 
        left: {
          style: 'thin',
          color: 'black'
        },
        right: {
          style: 'thin',
          color: 'black'
        },
        top: {
          style: 'thin',
          color: 'black'
        },
        bottom: {
          style: 'thin',
          color: 'black'
        }
      }
    };
    const Style_1={
      alignment:{
        vertical: 'center',
        wrapText: true,
      },
      font: {
        name: 'Times New Roman',
        size:10
      },
      border: { 
        left: {
          style: 'thin',
          color: 'black'
        },
        right: {
          style: 'thin',
          color: 'black'
        },
        top: {
          style: 'thin',
          color: 'black'
        },
        bottom: {
          style: 'thin',
          color: 'black'
        }
      }
    };
    let startRow = 10;
    let C=1;
    let startColumn = 1;
    Ok_disciplines.forEach(discipline=>{
      Object.keys(discipline).forEach((key,index)=>{
        worksheet.cell(startRow,startColumn).number(C).style(Style_);
        worksheet.cell(startRow,startColumn+1).string(discipline.conVK).style(Style_);
        worksheet.cell(startRow,startColumn+2).string(discipline.name).style(Style_1);
        worksheet.cell(startRow,startColumn+3).number(discipline.countCredit).style(Style_);
        worksheet.cell(startRow,startColumn+4).number(discipline.hoursInWeek*discipline.countWeek).style(Style_);
        worksheet.cell(startRow,startColumn+5).number(discipline.lectures+discipline.practicalLaboratory+discipline.seminar).style(Style_);
        worksheet.cell(startRow,startColumn+6).number(discipline.lectures).style(Style_);
        worksheet.cell(startRow,startColumn+7).number(discipline.practicalLaboratory).style(Style_);
        worksheet.cell(startRow,startColumn+8).number(discipline.seminar).style(Style_);
        worksheet.cell(startRow,startColumn+9).number(discipline.countCredit*30-discipline.hoursInWeek*discipline.countWeek).style(Style_);
        worksheet.cell(startRow,startColumn+11).string(discipline.formOfControl).style(Style_);
        worksheet.cell(startRow,startColumn+10).style(Style_);
        worksheet.cell(startRow,startColumn+12).style(Style_);
        worksheet.cell(startRow,startColumn+13).style(Style_);
        worksheet.cell(startRow,startColumn+14).style(Style_);
        worksheet.cell(startRow,startColumn+15).style(Style_);
        worksheet.cell(startRow,startColumn+16).style(Style_);
      });
      startRow++;
      C++;
    });
    const y=startRow-1;
    worksheet.cell(startRow,1).style(Style_);
    worksheet.cell(startRow,2).style(Style_);
    worksheet.cell(startRow,3).string("Всього:").style(Style_1);
    worksheet.cell(startRow,4).formula('SUM(D10:D'+y+')').style(Style_);
    worksheet.cell(startRow,5).formula('SUM(E10:E'+y+')').style(Style_);
    worksheet.cell(startRow,6).formula('SUM(F10:F'+y+')').style(Style_);
    worksheet.cell(startRow,7).formula('SUM(G10:G'+y+')').style(Style_);
    worksheet.cell(startRow,8).formula('SUM(H10:H'+y+')').style(Style_);
    worksheet.cell(startRow,9).formula('SUM(I10:I'+y+')').style(Style_);
    worksheet.cell(startRow,10).formula('SUM(J10:J'+y+')').style(Style_);
    worksheet.cell(startRow,11).formula('SUM(K10:K'+y+')').style(Style_);
    

    worksheet.cell(startRow,12).formula('COUNTIF(L10:L'+y+', "екзам")').style(Style_);
    worksheet.cell(startRow,13).style(Style_);
    worksheet.cell(startRow,14).style(Style_);
    worksheet.cell(startRow,15).style(Style_);
    worksheet.cell(startRow,16).style(Style_);
    worksheet.cell(startRow,17).style(Style_);
    startRow++;

    if (VK_disciplines.length === 0) {
      const D={
        alignment:{
          horizontal: 'center',
          vertical: 'center',
          wrapText: true,
        },
        font:{
          name:'TimeNewRoman'
        }
      };
      startRow++;
      worksheet.cell(startRow,1,startRow,3,true).string("Здобувач освіти__________").style(Style0);
      worksheet.cell(startRow,4,startRow,11,true).string("Відповідальна особа від закладу ФПО__________").style(Style0);
      worksheet.cell(startRow,13,startRow,16,true).string("Завідувач відділення_______________").style(Style0);
      worksheet.cell(startRow+1,3).string("(Підпис)").style(D);
      worksheet.cell(startRow+1,15,startRow+1,16,true).string("(Підпис)").style(D);
      worksheet.cell(startRow+1,17).string("(Підпис)").style(D);
    } 
    else {
      worksheet.cell(startRow,1,startRow,17,true).string("Вибіркові освітні компоненти").style(Style);
      startRow++;
      VK_disciplines.forEach(discipline=>{
        Object.keys(discipline).forEach((key,index)=>{
          worksheet.cell(startRow,startColumn).number(C).style(Style_);
          worksheet.cell(startRow,startColumn+1).string(discipline.conVK).style(Style_);
          worksheet.cell(startRow,startColumn+2).string(discipline.name).style(Style_1);
          worksheet.cell(startRow,startColumn+3).number(discipline.countCredit).style(Style_);
          worksheet.cell(startRow,startColumn+4).number(discipline.hoursInWeek*discipline.countWeek).style(Style_);
          worksheet.cell(startRow,startColumn+5).number(discipline.lectures+discipline.practicalLaboratory+discipline.seminar).style(Style_);
          worksheet.cell(startRow,startColumn+6).number(discipline.lectures).style(Style_);
          worksheet.cell(startRow,startColumn+7).number(discipline.practicalLaboratory).style(Style_);
          worksheet.cell(startRow,startColumn+8).number(discipline.seminar).style(Style_);
          worksheet.cell(startRow,startColumn+9).number(discipline.countCredit*30-discipline.hoursInWeek*discipline.countWeek).style(Style_);
          worksheet.cell(startRow,startColumn+11).string(discipline.formOfControl).style(Style_);
          worksheet.cell(startRow,startColumn+10).style(Style_);
          worksheet.cell(startRow,startColumn+12).style(Style_);
          worksheet.cell(startRow,startColumn+13).style(Style_);
          worksheet.cell(startRow,startColumn+14).style(Style_);
          worksheet.cell(startRow,startColumn+15).style(Style_);
          worksheet.cell(startRow,startColumn+16).style(Style_);
        });
        startRow++;
        C++;
      });
      const D={
        alignment:{
          horizontal: 'center',
          vertical: 'center',
          wrapText: true,
        },
        font:{
          name:'TimeNewRoman'
        }
      };
      let t=y+1;
      const x=startRow-1;
      worksheet.cell(startRow,1).style(Style_);
      worksheet.cell(startRow,2).style(Style_);
      worksheet.cell(startRow,3).string("Всього:").style(Style_1);
      worksheet.cell(startRow,4).formula('SUM(D'+t+':D'+x+')').style(Style_);
      worksheet.cell(startRow,5).formula('SUM(E'+t+':E'+x+')').style(Style_);
      worksheet.cell(startRow,6).formula('SUM(F'+t+':F'+x+')').style(Style_);
      worksheet.cell(startRow,7).formula('SUM(G'+t+':G'+x+')').style(Style_);
      worksheet.cell(startRow,8).formula('SUM(H'+t+':H'+x+')').style(Style_);
      worksheet.cell(startRow,9).formula('SUM(I'+t+':I'+x+')').style(Style_);
      worksheet.cell(startRow,10).formula('SUM(J'+t+':J'+x+')').style(Style_);
      worksheet.cell(startRow,11).formula('SUM(K'+t+':K'+x+')').style(Style_);
      worksheet.cell(startRow,12).formula('COUNTIF(L10:L'+x+', "екзам")').style(Style_);
      worksheet.cell(startRow,13).style(Style_);
      worksheet.cell(startRow,14).style(Style_);
      worksheet.cell(startRow,15).style(Style_);
      worksheet.cell(startRow,16).style(Style_);
      worksheet.cell(startRow,17).style(Style_);


      startRow++;
      startRow++;
      worksheet.cell(startRow,1,startRow,3,true).string("Здобувач освіти__________").style(Style0);
      worksheet.cell(startRow,4,startRow,11,true).string("Відповідальна особа від закладу ФПО__________").style(Style0);
      worksheet.cell(startRow,13,startRow,16,true).string("Завідувач відділення_______________").style(Style0);
      worksheet.cell(startRow+1,3).string("(Підпис)").style(D);
      worksheet.cell(startRow+1,15,startRow+1,16,true).string("(Підпис)").style(D);
      worksheet.cell(startRow+1,17).string("(Підпис)").style(D);
    }
  }

  function Practice_stydent(workbook,Practices){
    var worksheet = workbook.addWorksheet('Практика');

    const Style_0={
      alignment:{
        horizontal: 'center',
        vertical: 'center',
        wrapText: true,
      },
      font:{
        name:'TimeNewRoman',
        size:14,
        bold:true
      }
    };
    const Style_1={
      alignment:{
        horizontal: 'center',
        vertical: 'center',
        wrapText: true,
      },
      font:{
        name:'TimeNewRoman',
        size:12,
        bold:true
      },
      border: { 
        left: {
          style: 'thin',
          color: 'black'
        },
        right: {
          style: 'thin',
          color: 'black'
        },
        top: {
          style: 'thin',
          color: 'black'
        },
        bottom: {
          style: 'thin',
          color: 'black'
        }
      }
    };
    const Style_2={
      alignment:{
        horizontal: 'center',
        vertical: 'center',
        wrapText: true,
      },
      font:{
        name:'TimeNewRoman',
        size:12,
      },
      border: { 
        left: {
          style: 'thin',
          color: 'black'
        },
        right: {
          style: 'thin',
          color: 'black'
        },
        top: {
          style: 'thin',
          color: 'black'
        },
        bottom: {
          style: 'thin',
          color: 'black'
        }
      }
    };
    const Style_3={
      alignment:{
        vertical: 'center',
        wrapText: true,
      },
      font:{
        name:'TimeNewRoman',
        size:14,
      },
    };
    const Style_4={
      alignment:{
        horizontal: 'center',
        vertical: 'center',
        wrapText: true,
      },
      font:{
        name:'TimeNewRoman',
        size:12,
      },
      border: { 
        bottom: {
          style: 'thin',
          color: 'black'
        }
      }
    };
    const Style_5={
      alignment:{
        horizontal: 'center',
        vertical: 'center',
        wrapText: true,
      },
      font:{
        name:'TimeNewRoman',
        size:14,
      },
    };

    worksheet.cell(1,1,1,12,true).string("Практична підготовка").style(Style_0);
    worksheet.cell(3,1,4,1,true).string("№").style(Style_1);
    worksheet.column(1).setWidth(3);
    worksheet.cell(3,2,4,2,true).string("Код ОК").style(Style_1);
    worksheet.column(2).setWidth(6);
    worksheet.cell(3,3,4,3,true).string("Вид практики").style(Style_1);
    worksheet.column(3).setWidth(15);
    worksheet.cell(3,4,4,4,true).string("Кількість кредитів ЄКТС").style(Style_1);
    worksheet.cell(3,5,4,5,true).string("Курс").style(Style_1);
    worksheet.column(5).setWidth(6);
    worksheet.cell(3,6,4,6,true).string("Семестер").style(Style_1);
    worksheet.column(6).setWidth(10);
    worksheet.cell(3,7,3,9,true).string("Тривалість практики \n(дата)").style(Style_1);
    worksheet.cell(4,7).string("Від").style(Style_1);
    worksheet.cell(4,8).string("До").style(Style_1);
    worksheet.cell(4,9).string("Кількість тижнів").style(Style_1);
    worksheet.cell(3,10,4,10,true).string("Дата захисту").style(Style_1);
    worksheet.cell(3,11,3,12,true).string("Відмітка про виконання").style(Style_1);
    worksheet.cell(4,11).string("Оцінка за \nшкалою \n зпкладу \n освіти").style(Style_1);
    worksheet.cell(4,12).string("Прізвище \nвикладача").style(Style_1);


    let starRow=5;
    let number=1;
    Practices.forEach(practice => {
        Object.keys(practice).forEach((key,index)=>{
          worksheet.cell(starRow,1).number(number).style(Style_2);
          worksheet.cell(starRow,2).string(practice.conOK).style(Style_2);
          worksheet.cell(starRow,3).string(practice.name).style(Style_2);
          worksheet.cell(starRow,4).number(practice.countCredit).style(Style_2);
          worksheet.cell(starRow,5).number(practice.course).style(Style_2);
          worksheet.cell(starRow,6).number(practice.semester).style(Style_2);
          worksheet.cell(starRow,7).style(Style_2);
          worksheet.cell(starRow,8).style(Style_2);
          worksheet.cell(starRow,9).number(practice.lenght).style(Style_2);
          worksheet.cell(starRow,10).style(Style_2);
          worksheet.cell(starRow,11).style(Style_2);
          worksheet.cell(starRow,12).style(Style_2);
        });
        starRow++;
        number++;
    });

    starRow++;
    worksheet.cell(starRow,1,starRow,3,true).string("Завідувач відділення").style(Style_3);
    worksheet.cell(starRow,4,starRow,8,true).style(Style_4);
    starRow++;
    worksheet.cell(starRow,4,starRow,5,true).string("(Підпис").style(Style_5);
    worksheet.cell(starRow,6,starRow,8,true).string("Ініціали, прізвище)").style(Style_5);
  }
  
app.listen(3000, () => console.log('Сервер прослуховує порт 3000...'));














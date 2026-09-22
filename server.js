import "dotenv/config";
import express from "express";
import { Resend } from "resend";
import {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  WidthType, AlignmentType, BorderStyle, HeadingLevel
} from "docx";
import JSZip from "./zip-helper.js";
import path from "node:path";
import { fileURLToPath } from "node:url";

const app=express();
app.use(express.json({limit:"1mb"}));
app.use(express.static("public"));

const ILDAR = {
  name:"ИП Вруцкий Ильдар Анатольевич",
  inn:"772151379194",
  kpp:"773601001",
  ogrnip:"310774619600480",
  address:"109472, г. Москва",
  bank:"ПАО СБЕРБАНК г. МОСКВА",
  bik:"044525225",
  rs:"40802810040000032809",
  ks:"30101810400000000225",
  okpo:"00032537"
};

const borders={
  top:{style:BorderStyle.SINGLE,size:1,color:"000000"},
  bottom:{style:BorderStyle.SINGLE,size:1,color:"000000"},
  left:{style:BorderStyle.SINGLE,size:1,color:"000000"},
  right:{style:BorderStyle.SINGLE,size:1,color:"000000"},
  insideHorizontal:{style:BorderStyle.SINGLE,size:1,color:"000000"},
  insideVertical:{style:BorderStyle.SINGLE,size:1,color:"000000"}
};
const noBorders={
  top:{style:BorderStyle.NONE},bottom:{style:BorderStyle.NONE},
  left:{style:BorderStyle.NONE},right:{style:BorderStyle.NONE},
  insideHorizontal:{style:BorderStyle.NONE},insideVertical:{style:BorderStyle.NONE}
};

function ruDate(v){
  if(!v) return "";
  const d=new Date(v+"T00:00:00");
  const months=["января","февраля","марта","апреля","мая","июня","июля","августа","сентября","октября","ноября","декабря"];
  return `${String(d.getDate()).padStart(2,"0")} ${months[d.getMonth()]} ${d.getFullYear()} г.`;
}
function fmt(n){
  return new Intl.NumberFormat("ru-RU",{minimumFractionDigits:2,maximumFractionDigits:2}).format(Number(n||0));
}
function totalOf(items=[]){
  return items.reduce((s,x)=>s+(Number(x.qty)||0)*(Number(x.price)||0),0);
}
function text(n,bold=false,size=20){
  return new TextRun({text:String(n??""),bold,size,font:"Times New Roman"});
}
function p(content="",opts={}){
  const children=Array.isArray(content)?content:[text(content,!!opts.bold,opts.size||20)];
  return new Paragraph({
    children,
    alignment:opts.align||AlignmentType.JUSTIFIED,
    spacing:{before:opts.before??0,after:opts.after??60,line:opts.line??260},
    keepNext:opts.keepNext||false,
    keepLines:opts.keepLines||false,
    heading:opts.heading
  });
}
function heading(t){
  return new Paragraph({
    children:[text(t,true,22)],
    alignment:AlignmentType.CENTER,
    spacing:{before:180,after:90},
    keepNext:true
  });
}
function cell(content,opts={}){
  const paras=Array.isArray(content)?content:[p(content,{align:opts.align||AlignmentType.LEFT,size:opts.size||18})];
  return new TableCell({children:paras,width:opts.width?{size:opts.width,type:WidthType.PERCENTAGE}:undefined});
}

function makeInvoice(data){
  const items=data.items||[];
  const total=totalOf(items);
  const pageWidth=11906;
  const contentWidth=10206;

  const fixedCell=(content,width,opts={})=>{
    const paras=Array.isArray(content)?content:[p(content,{
      align:opts.align||AlignmentType.LEFT,
      size:opts.size||18,
      after:opts.after??0,
      line:opts.line??220
    })];
    return new TableCell({
      children:paras,
      width:{size:width,type:WidthType.DXA},
      margins:{top:70,bottom:70,left:90,right:90},
      shading:opts.shading?{fill:opts.shading}:undefined,
      verticalAlign:"center"
    });
  };

  const bankTable=new Table({
    width:{size:contentWidth,type:WidthType.DXA},
    columnWidths:[6100,900,3206],
    borders,
    rows:[
      new TableRow({children:[
        fixedCell(ILDAR.bank,6100,{size:17}),
        fixedCell("БИК",900,{size:17}),
        fixedCell(ILDAR.bik,3206,{size:17})
      ]}),
      new TableRow({children:[
        fixedCell("Банк получателя",6100,{size:15}),
        fixedCell("Сч. №",900,{size:15}),
        fixedCell(ILDAR.ks,3206,{size:15})
      ]}),
      new TableRow({children:[
        fixedCell(`ИНН ${ILDAR.inn}     КПП ${ILDAR.kpp}`,6100,{size:15}),
        fixedCell("Сч. №",900,{size:15}),
        fixedCell(ILDAR.rs,3206,{size:15})
      ]}),
      new TableRow({children:[
        fixedCell([p([text(ILDAR.name,false,16),text("   Получатель",false,14)],{align:AlignmentType.LEFT,after:0,line:220})],10206)
      ]})
    ]
  });

  const buyerLines=[
    data.customerShort||"",
    data.customerFull?`Полное наименование: ${data.customerFull}`:"",
    data.customerAddress?`Юридический адрес: ${data.customerAddress}`:"",
    [data.customerInn&&`ИНН ${data.customerInn}`,data.customerKpp&&`КПП ${data.customerKpp}`,data.customerOgrn&&`ОГРН ${data.customerOgrn}`].filter(Boolean).join(", "),
    [data.customerOkpo&&`ОКПО ${data.customerOkpo}`,data.customerOkved&&`ОКВЭД ${data.customerOkved}`,data.customerOkato&&`ОКАТО ${data.customerOkato}`,data.customerOktmo&&`ОКТМО ${data.customerOktmo}`].filter(Boolean).join(", "),
    data.customerRs?`р/с ${data.customerRs}${data.customerBank?` в ${data.customerBank}`:""}`:"",
    [data.customerKs&&`к/с ${data.customerKs}`,data.customerBik&&`БИК ${data.customerBik}`].filter(Boolean).join(", ")
  ].filter(Boolean);

  const itemWidths=[500,5200,800,700,1450,1556];
  const headerLabels=["№","Товары (работы, услуги)","Кол-во","Ед.","Цена","Сумма"];
  const rows=[
    new TableRow({children:headerLabels.map((h,i)=>fixedCell(h,itemWidths[i],{
      size:16,bold:true,align:AlignmentType.CENTER,shading:"F1F3F5"
    }))})
  ];

  items.forEach((it,idx)=>{
    const sum=(Number(it.qty)||0)*(Number(it.price)||0);
    rows.push(new TableRow({children:[
      fixedCell(String(idx+1),itemWidths[0],{align:AlignmentType.CENTER,size:16}),
      fixedCell(it.description||"",itemWidths[1],{size:16}),
      fixedCell(String(it.qty||0),itemWidths[2],{align:AlignmentType.CENTER,size:16}),
      fixedCell(it.unit||"шт",itemWidths[3],{align:AlignmentType.CENTER,size:16}),
      fixedCell(fmt(it.price),itemWidths[4],{align:AlignmentType.RIGHT,size:16}),
      fixedCell(fmt(sum),itemWidths[5],{align:AlignmentType.RIGHT,size:16})
    ]}));
  });

  const itemsTable=new Table({
    width:{size:contentWidth,type:WidthType.DXA},
    columnWidths:itemWidths,
    borders,
    rows
  });

  const totalsTable=new Table({
    width:{size:5200,type:WidthType.DXA},
    columnWidths:[3200,2000],
    alignment:AlignmentType.RIGHT,
    borders:noBorders,
    rows:[
      new TableRow({children:[
        fixedCell("Итого:",3200,{align:AlignmentType.RIGHT,size:18}),
        fixedCell(fmt(total),2000,{align:AlignmentType.RIGHT,size:18})
      ]}),
      new TableRow({children:[
        fixedCell("В том числе НДС:",3200,{align:AlignmentType.RIGHT,size:18}),
        fixedCell("",2000,{align:AlignmentType.RIGHT,size:18})
      ]}),
      new TableRow({children:[
        fixedCell("Всего к оплате:",3200,{align:AlignmentType.RIGHT,size:19}),
        fixedCell(fmt(total),2000,{align:AlignmentType.RIGHT,size:19})
      ]})
    ]
  });

  return new Document({sections:[{
    properties:{
      page:{
        size:{width:pageWidth,height:16838},
        margin:{top:650,right:850,bottom:650,left:850}
      }
    },
    children:[
      bankTable,
      p("",{after:90}),
      p([text(`Счет на оплату №${data.invoiceNumber||""} от ${ruDate(data.invoiceDate)}`,true,28)],{align:AlignmentType.LEFT,after:100,line:260}),
      p([text("Поставщик: ",false,17),text(`${ILDAR.name}, ИНН ${ILDAR.inn}, ${ILDAR.address}`,true,17)],{align:AlignmentType.LEFT,after:70,line:230}),
      p([text("Покупатель: ",false,17),text(buyerLines[0]||"",true,17)],{align:AlignmentType.LEFT,after:20,line:230}),
      ...buyerLines.slice(1).map(x=>p([text("                 "+x,false,16)],{align:AlignmentType.LEFT,after:15,line:220})),
      p("",{after:90}),
      itemsTable,
      p("",{after:70}),
      totalsTable,
      p("",{after:70}),
      p([text(`Всего наименований ${items.length}, на сумму ${fmt(total)} руб.`,false,17)],{align:AlignmentType.LEFT,after:35,line:230}),
      p([text(`Сумма прописью: ${numberToWordsRub(total)}`,true,18)],{align:AlignmentType.LEFT,after:140,line:230}),
      new Table({
        width:{size:contentWidth,type:WidthType.DXA},
        columnWidths:[2600,2700,2700,2206],
        borders:noBorders,
        rows:[
          new TableRow({children:[
            fixedCell("Руководитель",2600,{size:16}),
            fixedCell("____________________",2700,{align:AlignmentType.CENTER,size:16}),
            fixedCell("____________________",2700,{align:AlignmentType.CENTER,size:16}),
            fixedCell("Вруцкий И.А.",2206,{align:AlignmentType.CENTER,size:16})
          ]}),
          new TableRow({children:[
            fixedCell("Главный (старший) бухгалтер",2600,{size:16}),
            fixedCell("",2700,{size:16}),
            fixedCell("____________________",2700,{align:AlignmentType.CENTER,size:16}),
            fixedCell("Вруцкий И.А.",2206,{align:AlignmentType.CENTER,size:16})
          ]})
        ]
      })
    ]
  }]});
}

function makeContract(data){
  const items=data.items||[];
  const total=totalOf(items);
  const adv=total*(Number(data.advancePercent)||0)/100;
  const bal=total-adv;
  const contentWidth=10206;

  const workList=items.map((it,idx)=>
    p(`${idx+1}) ${it.description||""}; количество: ${it.qty||0} ${it.unit||"шт"}; цена: ${fmt(it.price)} руб.; сумма: ${fmt((it.qty||0)*(it.price||0))} руб.`,{
      align:AlignmentType.JUSTIFIED,after:35,line:255
    })
  );

  const customerReq=[
    data.customerShort||"",
    data.customerAddress?`Юр. адрес: ${data.customerAddress}`:"",
    [data.customerInn&&`ИНН ${data.customerInn}`,data.customerKpp&&`КПП ${data.customerKpp}`].filter(Boolean).join(" / "),
    data.customerOgrn?`ОГРН: ${data.customerOgrn}`:"",
    data.customerRs?`р/с: ${data.customerRs}`:"",
    data.customerBank?`Банк: ${data.customerBank}`:"",
    data.customerKs?`к/с: ${data.customerKs}`:"",
    data.customerBik?`БИК: ${data.customerBik}`:""
  ].filter(Boolean).map(x=>p(x,{align:AlignmentType.LEFT,size:17,after:25,line:220}));

  const intro=`${ILDAR.name}, действующий на основании регистрации в качестве индивидуального предпринимателя, ИНН ${ILDAR.inn}, ОГРНИП ${ILDAR.ogrnip}, именуемый в дальнейшем «Исполнитель», с одной стороны, и ${data.customerShort||"____________________________"}${data.representative?`, в лице ${data.representativeTitle||""} ${data.representative}`:""}, именуемый в дальнейшем «Заказчик», с другой стороны, а совместно именуемые «Стороны», заключили настоящий договор о нижеследующем:`;

  return new Document({sections:[{
    properties:{
      page:{
        size:{width:11906,height:16838},
        margin:{top:700,right:850,bottom:700,left:850}
      }
    },
    children:[
      p([text(`ДОГОВОР № ${data.contractNumber||"________"}`,true,27)],{align:AlignmentType.CENTER,after:80,line:260}),
      new Table({
        width:{size:contentWidth,type:WidthType.DXA},
        columnWidths:[5103,5103],
        borders:noBorders,
        rows:[new TableRow({children:[
          new TableCell({children:[p("г. Москва",{align:AlignmentType.LEFT,size:18,after:0})],width:{size:5103,type:WidthType.DXA}}),
          new TableCell({children:[p(ruDate(data.contractDate),{align:AlignmentType.RIGHT,size:18,after:0})],width:{size:5103,type:WidthType.DXA}})
        ]})]
      }),
      p("",{after:45}),
      p(intro,{after:110,line:270}),
      heading("1. ПРЕДМЕТ ДОГОВОРА"),
      p("1.1. Заказчик поручает, а Исполнитель принимает на себя обязательство по изготовлению, покраске, доставке и монтажу металлических конструкций.",{after:45}),
      ...workList,
      p(`1.2. Объект работ: ${data.workAddress||"____________________________________________"}.`,{after:45}),
      p("1.3. Заказчик принимает и оплачивает выполненные Исполнителем работы.",{after:70}),
      heading("2. ОБЯЗАТЕЛЬСТВА СТОРОН"),
      p("2.1. Заказчик обязуется:",{after:35}),
      p("2.1.1. Оплатить аванс в размере, установленном настоящим договором.",{after:35}),
      p("2.1.2. По завершении работ подписать Акт приемки-сдачи и оплатить работы в соответствии с договором.",{after:35}),
      p("2.1.3. Предоставить Исполнителю подготовленную и безопасную площадку со свободным доступом, источником электропитания и согласовать время монтажных работ.",{after:45}),
      p("2.2. Исполнитель обязуется:",{after:35}),
      p("2.2.1. Выполнить лично и своими материалами все работы в сроки, указанные в п. 5.1 настоящего договора.",{after:35}),
      p("2.2.2. Обеспечить надлежащее качество выполненных работ.",{after:70}),
      heading("3. ЦЕНА ДОГОВОРА И ПОРЯДОК РАСЧЕТОВ"),
      p(`3.1. Цена настоящего договора составляет ${fmt(total)} руб.`,{after:35}),
      p(`3.2. Заказчик выплачивает Исполнителю аванс в размере ${data.advancePercent||0}% — ${fmt(adv)} руб. для приобретения материалов, изготовления и доставки.`,{after:35}),
      p(`3.3. Окончательная оплата оставшейся суммы ${fmt(bal)} руб. производится в день приема монтажа и подписания Акта приемки-сдачи.`,{after:35}),
      p("3.4. Оплата производится путем безналичного платежа на расчетный счет Исполнителя.",{after:35}),
      p("3.5. В случае задержки окончательной оплаты Заказчик выплачивает Исполнителю неустойку в размере 0,05% от суммы задолженности за каждый день просрочки платежа.",{after:35}),
      p("3.6. Дополнительные виды работ, выявившиеся в процессе выполнения работ, оформляются актами и дополнительными сметами.",{after:70}),
      heading("4. ОТВЕТСТВЕННОСТЬ СТОРОН И ПОРЯДОК РАЗРЕШЕНИЯ СПОРОВ"),
      p("4.1. Стороны несут ответственность за неисполнение или ненадлежащее исполнение своих обязательств по настоящему Договору.",{after:35}),
      p("4.2. Меры ответственности сторон, не предусмотренные в настоящем договоре, применяются в соответствии с законодательством Российской Федерации.",{after:35}),
      p("4.3. Спорные ситуации разрешаются путем переговоров между Заказчиком и Исполнителем.",{after:70}),
      heading("5. СРОКИ ИСПОЛНЕНИЯ РАБОТ"),
      p(`5.1. Изготовление металлических конструкций должно быть выполнено Исполнителем в течение ${data.workDays||40} рабочих дней с момента поступления аванса на расчетный счет Исполнителя.`,{after:35}),
      p("5.2. В случае задержки сдачи работ по настоящему договору Исполнитель выплачивает Заказчику неустойку в размере 0,05% от общей стоимости работ за каждый день просрочки.",{after:70}),
      heading("6. УСЛОВИЯ ИЗМЕНЕНИЯ И РАСТОРЖЕНИЯ ДОГОВОРА"),
      p("6.1. Настоящий договор может быть расторгнут по соглашению Сторон либо в иных случаях, предусмотренных законодательством Российской Федерации.",{after:35}),
      p("6.2. Любые изменения и дополнения действительны, если они совершены в письменной форме и подписаны обеими Сторонами.",{after:70}),
      heading("7. ГАРАНТИЙНЫЕ ОБЯЗАТЕЛЬСТВА"),
      p(`7.1. Гарантийный срок на выполненные работы составляет ${data.warrantyMonths||12} месяцев.`,{after:35}),
      p("7.2. Гарантия не распространяется на механические повреждения, возникшие после приемки по вине Заказчика или третьих лиц, а также вследствие нарушения правил эксплуатации.",{after:70}),
      heading("8. ВСТУПЛЕНИЕ В СИЛУ ДОГОВОРА"),
      p("8.1. Договор вступает в силу с момента его подписания и действует до полного исполнения сторонами всех обязательств.",{after:35}),
      p("8.2. Настоящий договор составлен в двух экземплярах, имеющих одинаковую юридическую силу.",{after:110}),
      heading("АДРЕСА И БАНКОВСКИЕ РЕКВИЗИТЫ СТОРОН"),
      new Table({
        width:{size:contentWidth,type:WidthType.DXA},
        columnWidths:[5103,5103],
        borders:noBorders,
        rows:[new TableRow({children:[
          new TableCell({
            width:{size:5103,type:WidthType.DXA},
            margins:{top:80,bottom:80,left:80,right:180},
            children:[
              p([text("Исполнитель:",true,18)],{align:AlignmentType.LEFT,after:50}),
              p(ILDAR.name,{align:AlignmentType.LEFT,size:17,after:25}),
              p(`ИНН ${ILDAR.inn}`,{align:AlignmentType.LEFT,size:17,after:25}),
              p(`ОГРНИП ${ILDAR.ogrnip}`,{align:AlignmentType.LEFT,size:17,after:25}),
              p(`Адрес: ${ILDAR.address}`,{align:AlignmentType.LEFT,size:17,after:25}),
              p(`р/с ${ILDAR.rs}`,{align:AlignmentType.LEFT,size:17,after:25}),
              p(`Банк: ${ILDAR.bank}`,{align:AlignmentType.LEFT,size:17,after:25}),
              p(`к/с ${ILDAR.ks}`,{align:AlignmentType.LEFT,size:17,after:25}),
              p(`БИК ${ILDAR.bik}`,{align:AlignmentType.LEFT,size:17,after:100}),
              p("________________ / Вруцкий И.А.",{align:AlignmentType.LEFT,size:17,after:0})
            ]
          }),
          new TableCell({
            width:{size:5103,type:WidthType.DXA},
            margins:{top:80,bottom:80,left:180,right:80},
            children:[
              p([text("Заказчик:",true,18)],{align:AlignmentType.LEFT,after:50}),
              ...customerReq,
              p("",{align:AlignmentType.LEFT,after:80}),
              p("________________ / __________________",{align:AlignmentType.LEFT,size:17,after:0})
            ]
          })
        ]})]
      })
    ]
  }]});
}

function numberToWordsRub(amount){
  const n=Math.round(Number(amount||0));
  if(!Number.isFinite(n)) return "";
  return `${new Intl.NumberFormat("ru-RU").format(n)} рублей 00 копеек`;
}

async function buildFiles(data){
  if(!data.customerShort) throw new Error("Введите наименование Заказчика.");
  if(!Array.isArray(data.items) || !data.items.length) throw new Error("Добавьте хотя бы одну позицию.");
  const invoice=await Packer.toBuffer(makeInvoice(data));
  const contract=await Packer.toBuffer(makeContract(data));
  return {invoice,contract};
}

app.post("/api/generate",async(req,res)=>{
  try{
    const {invoice,contract}=await buildFiles(req.body);
    const zip=new JSZip();
    zip.file(`Счет_${req.body.invoiceNumber||""}.docx`,invoice);
    zip.file(`Договор_${req.body.contractNumber||""}.docx`,contract);
    const buf=await zip.generateAsync({type:"nodebuffer"});
    res.setHeader("Content-Type","application/zip");
    res.setHeader("Content-Disposition",'attachment; filename="documents.zip"');
    res.send(buf);
  }catch(e){
    res.status(400).json({error:e.message});
  }
});

app.post("/api/send",async(req,res)=>{
  try{
    if(!req.body.email) throw new Error("Введите e-mail.");
    if(!process.env.RESEND_API_KEY) throw new Error("Отправка e-mail не настроена. Добавьте RESEND_API_KEY в файл .env.");
    const {invoice,contract}=await buildFiles(req.body);
    const resend=new Resend(process.env.RESEND_API_KEY);
    await resend.emails.send({
      from:process.env.FROM_EMAIL||"Документы <onboarding@resend.dev>",
      to:req.body.email,
      subject:`Счёт №${req.body.invoiceNumber||""} и договор №${req.body.contractNumber||""}`,
      html:`<p>Во вложении счёт и договор.</p><p>${ILDAR.name}</p>`,
      attachments:[
        {filename:`Счет_${req.body.invoiceNumber||""}.docx`,content:invoice.toString("base64")},
        {filename:`Договор_${req.body.contractNumber||""}.docx`,content:contract.toString("base64")}
      ]
    });
    res.json({message:`Документы отправлены на ${req.body.email}.`});
  }catch(e){
    res.status(400).json({error:e.message});
  }
});

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

app.get("/", (req, res) => {
  res.sendFile(path.join(__dirname, "public", "index.html"));
});

export default app;

if (!process.env.VERCEL) {
  const port = process.env.PORT || 3000;
  app.listen(port, () => console.log(`http://localhost:${port}`));
}

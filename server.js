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
    spacing:{after:opts.after??0,line:opts.line??240},
    heading:opts.heading
  });
}
function heading(t){
  return new Paragraph({
    children:[text(t,true,22)],
    alignment:AlignmentType.CENTER,
    spacing:{before:120,after:60}
  });
}
function cell(content,opts={}){
  const paras=Array.isArray(content)?content:[p(content,{align:opts.align||AlignmentType.LEFT,size:opts.size||18})];
  return new TableCell({children:paras,width:opts.width?{size:opts.width,type:WidthType.PERCENTAGE}:undefined});
}

function makeInvoice(data){
  const items=data.items||[];
  const total=totalOf(items);

  const bankTable=new Table({
    width:{size:100,type:WidthType.PERCENTAGE},
    borders,
    rows:[
      new TableRow({children:[
        cell([p([text(ILDAR.bank,false,18)])],{width:55}),
        cell([p([text("БИК",false,18)])],{width:10}),
        cell([p([text(ILDAR.bik,false,18)])],{width:35})
      ]}),
      new TableRow({children:[
        cell([p([text("Банк получателя",false,16)])],{width:55}),
        cell([p([text("Сч. №",false,16)])],{width:10}),
        cell([p([text(ILDAR.ks,false,16)])],{width:35})
      ]}),
      new TableRow({children:[
        cell([p([text(`ИНН ${ILDAR.inn}   КПП ${ILDAR.kpp}`,false,16)])],{width:55}),
        cell([p([text("Сч. №",false,16)])],{width:10}),
        cell([p([text(ILDAR.rs,false,16)])],{width:35})
      ]}),
      new TableRow({children:[
        cell([p([text(ILDAR.name,false,16),text("\nПолучатель",false,14)])],{width:100})
      ]})
    ]
  });

  const rows=[
    new TableRow({children:["№","Товары (работы, услуги)","Кол-во","Ед.","Цена","Сумма"].map((h,i)=>cell([p([text(h,true,16)],{align:AlignmentType.CENTER})],{width:[5,55,8,7,12,13][i]}))})
  ];
  items.forEach((it,idx)=>{
    rows.push(new TableRow({children:[
      cell(String(idx+1),{width:5,align:AlignmentType.CENTER}),
      cell(it.description||"",{width:55}),
      cell(String(it.qty||0),{width:8,align:AlignmentType.CENTER}),
      cell(it.unit||"шт",{width:7,align:AlignmentType.CENTER}),
      cell(fmt(it.price),{width:12,align:AlignmentType.CENTER}),
      cell(fmt((it.qty||0)*(it.price||0)),{width:13,align:AlignmentType.CENTER})
    ]}));
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

  return new Document({sections:[{properties:{page:{margin:{top:600,right:700,bottom:600,left:700}}},children:[
    bankTable,
    p("",{after:40}),
    p([text(`Счет на оплату №${data.invoiceNumber||""} от ${ruDate(data.invoiceDate)}`,true,28)],{align:AlignmentType.LEFT,after:80}),
    p([text("Поставщик: ",false,18),text(`${ILDAR.name}, ИНН ${ILDAR.inn}, ${ILDAR.address}`,true,18)],{align:AlignmentType.LEFT,after:80}),
    p([text("Покупатель: ",false,18),text(buyerLines[0]||"",true,18)],{align:AlignmentType.LEFT}),
    ...buyerLines.slice(1).map(x=>p([text("              "+x,false,16)],{align:AlignmentType.LEFT})),
    p("",{after:80}),
    new Table({width:{size:100,type:WidthType.PERCENTAGE},borders,rows}),
    p("",{after:50}),
    p([text(`Итого: ${fmt(total)}`,true,20)],{align:AlignmentType.RIGHT}),
    p([text("В том числе НДС: ",true,20)],{align:AlignmentType.RIGHT}),
    p([text(`Всего к оплате: ${fmt(total)}`,true,20)],{align:AlignmentType.RIGHT}),
    p("",{after:40}),
    p([text(`Всего наименований ${items.length}, на сумму ${fmt(total)} руб.`,false,18)],{align:AlignmentType.LEFT}),
    p([text(`Сумма прописью: ${numberToWordsRub(total)}`,true,20)],{align:AlignmentType.LEFT,after:100}),
    p([text("Руководитель ____________________     ____________________     Вруцкий И.А",false,18)],{align:AlignmentType.LEFT}),
    p([text("Главный (старший) бухгалтер          ____________________     Вруцкий И.А",false,18)],{align:AlignmentType.LEFT})
  ]}]});
}

function makeContract(data){
  const items=data.items||[];
  const total=totalOf(items);
  const adv=total*(Number(data.advancePercent)||0)/100;
  const bal=total-adv;

  const workList=items.map((it,idx)=>
    p(`${idx+1}) ${it.description||""}; количество: ${it.qty||0} ${it.unit||"шт"}; цена: ${fmt(it.price)} руб.; сумма: ${fmt((it.qty||0)*(it.price||0))} руб.`,{align:AlignmentType.JUSTIFIED})
  );

  const customerReq=[
    data.customerShort||"",
    data.customerAddress?`Юр. адрес: ${data.customerAddress}`:"",
    [data.customerInn&&`ИНН ${data.customerInn}`,data.customerKpp&&`КПП ${data.customerKpp}`].filter(Boolean).join("/"),
    data.customerOgrn?`ОГРН: ${data.customerOgrn}`:"",
    data.customerRs?`р/с: ${data.customerRs}`:"",
    data.customerBank?`в ${data.customerBank}`:"",
    data.customerKs?`к/с: ${data.customerKs}`:"",
    data.customerBik?`БИК: ${data.customerBik}`:""
  ].filter(Boolean).map(x=>p(x,{align:AlignmentType.LEFT,size:18}));

  const intro=`${ILDAR.name}, действующий на основании регистрации в качестве индивидуального предпринимателя, ИНН ${ILDAR.inn}, ОГРНИП ${ILDAR.ogrnip}, именуемый в дальнейшем «Исполнитель», с одной стороны, и ${data.customerShort||"____________________________"}${data.representative?`, в лице ${data.representativeTitle||""} ${data.representative}`:""}, именуемый в дальнейшем «Заказчик», с другой стороны, а совместно именуемые «Стороны», заключили настоящий договор о нижеследующем:`;

  return new Document({sections:[{properties:{page:{margin:{top:700,right:850,bottom:700,left:850}}},children:[
    p([text(`ДОГОВОР № ${data.contractNumber||"________"}`,true,26)],{align:AlignmentType.CENTER,after:40}),
    new Table({width:{size:100,type:WidthType.PERCENTAGE},borders:noBorders,rows:[
      new TableRow({children:[
        cell([p([text("г. Москва",false,18)],{align:AlignmentType.LEFT})],{width:50}),
        cell([p([text(ruDate(data.contractDate),false,18)],{align:AlignmentType.RIGHT})],{width:50})
      ]})
    ]}),
    p(intro,{after:80}),
    heading("1. ПРЕДМЕТ ДОГОВОРА"),
    p("1.1. Заказчик поручает, а Исполнитель принимает на себя обязательство по изготовлению, покраске, доставке и монтажу металлических конструкций."),
    ...workList,
    p(`1.2. Объект работ: ${data.workAddress||"____________________________________________"}.`),
    p("1.3. Заказчик принимает и оплачивает выполненные Исполнителем работы."),
    heading("2. ОБЯЗАТЕЛЬСТВА СТОРОН"),
    p("2.1. Заказчик обязуется:"),
    p("2.1.1. Оплатить аванс в размере, установленном настоящим договором."),
    p("2.1.2. По завершении работ подписать Акт приемки-сдачи и оплатить работы в соответствии с договором."),
    p("2.1.3. Предоставить Исполнителю подготовленную и безопасную площадку со свободным доступом, источником электропитания и согласовать время монтажных работ."),
    p("2.2. Исполнитель обязуется:"),
    p("2.2.1. Выполнить лично и своими материалами все работы в сроки, указанные в п. 5.1 настоящего договора."),
    p("2.2.2. Обеспечить надлежащее качество выполненных работ."),
    heading("3. ЦЕНА ДОГОВОРА И ПОРЯДОК РАСЧЕТОВ"),
    p(`3.1. Цена настоящего договора составляет ${fmt(total)} руб.`),
    p(`3.2. Заказчик выплачивает Исполнителю аванс в размере ${data.advancePercent||0}% — ${fmt(adv)} руб. для приобретения материалов, изготовления и доставки.`),
    p(`3.3. Окончательная оплата оставшейся суммы ${fmt(bal)} руб. производится в день приема монтажа и подписания Акта приемки-сдачи.`),
    p("3.4. Оплата производится путем безналичного платежа на расчетный счет Исполнителя."),
    p("3.5. В случае задержки окончательной оплаты Заказчик выплачивает Исполнителю неустойку в размере 0,05% от суммы задолженности за каждый день просрочки платежа."),
    p("3.6. Дополнительные виды работ, выявившиеся в процессе выполнения работ, оформляются актами и дополнительными сметами."),
    heading("4. ОТВЕТСТВЕННОСТЬ СТОРОН И ПОРЯДОК РАЗРЕШЕНИЯ СПОРОВ"),
    p("4.1. Стороны несут ответственность за неисполнение или ненадлежащее исполнение своих обязательств по настоящему Договору."),
    p("4.2. Меры ответственности сторон, не предусмотренные в настоящем договоре, применяются в соответствии с законодательством Российской Федерации."),
    p("4.3. Спорные ситуации разрешаются путем переговоров между Заказчиком и Исполнителем."),
    heading("5. СРОКИ ИСПОЛНЕНИЯ РАБОТ"),
    p(`5.1. Изготовление металлических конструкций должно быть выполнено Исполнителем в течение ${data.workDays||40} рабочих дней с момента поступления аванса на расчетный счет Исполнителя.`),
    p("5.2. В случае задержки сдачи работ по настоящему договору Исполнитель выплачивает Заказчику неустойку в размере 0,05% от общей стоимости работ за каждый день просрочки."),
    heading("6. УСЛОВИЯ ИЗМЕНЕНИЯ И РАСТОРЖЕНИЯ ДОГОВОРА"),
    p("6.1. Настоящий договор может быть расторгнут по соглашению Сторон либо в иных случаях, предусмотренных законодательством Российской Федерации."),
    p("6.2. Любые изменения и дополнения действительны, если они совершены в письменной форме и подписаны обеими Сторонами."),
    heading("7. ГАРАНТИЙНЫЕ ОБЯЗАТЕЛЬСТВА"),
    p(`7.1. Гарантийный срок на выполненные работы составляет ${data.warrantyMonths||12} месяцев.`),
    p("7.2. Гарантия не распространяется на механические повреждения, возникшие после приемки по вине Заказчика или третьих лиц, а также вследствие нарушения правил эксплуатации."),
    heading("8. ВСТУПЛЕНИЕ В СИЛУ ДОГОВОРА"),
    p("8.1. Договор вступает в силу с момента его подписания и действует до полного исполнения сторонами всех обязательств."),
    p("8.2. Настоящий договор составлен в двух экземплярах, имеющих одинаковую юридическую силу."),
    heading("АДРЕСА И БАНКОВСКИЕ РЕКВИЗИТЫ СТОРОН"),
    new Table({width:{size:100,type:WidthType.PERCENTAGE},borders:noBorders,rows:[
      new TableRow({children:[
        cell([
          p([text("Исполнитель:",true,18)],{align:AlignmentType.LEFT}),
          p(ILDAR.name,{align:AlignmentType.LEFT,size:18}),
          p(`ИНН ${ILDAR.inn}`,{align:AlignmentType.LEFT,size:18}),
          p(`ОГРНИП ${ILDAR.ogrnip}`,{align:AlignmentType.LEFT,size:18}),
          p(`Адрес: ${ILDAR.address}`,{align:AlignmentType.LEFT,size:18}),
          p(`р/с ${ILDAR.rs}`,{align:AlignmentType.LEFT,size:18}),
          p(`в ${ILDAR.bank}`,{align:AlignmentType.LEFT,size:18}),
          p(`к/с ${ILDAR.ks}`,{align:AlignmentType.LEFT,size:18}),
          p(`БИК ${ILDAR.bik}`,{align:AlignmentType.LEFT,size:18}),
          p("",{align:AlignmentType.LEFT}),
          p("________________ / Вруцкий И.А.",{align:AlignmentType.LEFT,size:18})
        ],{width:50}),
        cell([
          p([text("Заказчик:",true,18)],{align:AlignmentType.LEFT}),
          ...customerReq,
          p("",{align:AlignmentType.LEFT}),
          p("________________ / __________________",{align:AlignmentType.LEFT,size:18})
        ],{width:50})
      ]})
    ]})
  ]}]});
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

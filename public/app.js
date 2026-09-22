const form = document.getElementById("form");
const itemsEl = document.getElementById("items");
const tpl = document.getElementById("itemTemplate");
const statusEl = document.getElementById("status");

function todayISO(){
  const d=new Date();
  const z=n=>String(n).padStart(2,"0");
  return `${d.getFullYear()}-${z(d.getMonth()+1)}-${z(d.getDate())}`;
}
form.contractDate.value = todayISO();
form.invoiceDate.value = todayISO();

function money(n){
  return new Intl.NumberFormat("ru-RU",{style:"currency",currency:"RUB",maximumFractionDigits:2}).format(Number(n||0));
}
function recalc(){
  const rows=[...document.querySelectorAll(".item-row")];
  const total=rows.reduce((s,row)=>s+(+row.querySelector(".itemQty").value||0)*(+row.querySelector(".itemPrice").value||0),0);
  const pct=+form.advancePercent.value||0;
  const advance=total*pct/100;
  document.getElementById("total").textContent=money(total);
  document.getElementById("advance").textContent=money(advance);
  document.getElementById("balance").textContent=money(total-advance);
}
function addItem(data={}){
  const node=tpl.content.cloneNode(true);
  const row=node.querySelector(".item-row");
  row.querySelector(".itemDescription").value=data.description||"";
  row.querySelector(".itemQty").value=data.qty ?? 1;
  row.querySelector(".itemUnit").value=data.unit||"шт";
  row.querySelector(".itemPrice").value=data.price ?? 0;
  row.addEventListener("input",recalc);
  row.querySelector(".remove").addEventListener("click",()=>{row.remove();recalc()});
  itemsEl.appendChild(node);
  recalc();
}
addItem({description:"Изготовление и установка кованой металлической конструкции — оконная решетка, размер 2050×820 мм",qty:1,unit:"шт",price:120000});
addItem({description:"Изготовление и установка кованой металлической конструкции — заборные ограждения, размер 1200×570×450 мм",qty:6,unit:"шт",price:120000});
document.getElementById("addItem").addEventListener("click",()=>addItem());
form.advancePercent.addEventListener("input",recalc);

function payload(){
  const fd=new FormData(form);
  const data=Object.fromEntries(fd.entries());
  data.items=[...document.querySelectorAll(".item-row")].map(row=>({
    description:row.querySelector(".itemDescription").value,
    qty:Number(row.querySelector(".itemQty").value||0),
    unit:row.querySelector(".itemUnit").value,
    price:Number(row.querySelector(".itemPrice").value||0)
  }));
  data.advancePercent=Number(data.advancePercent||0);
  data.workDays=Number(data.workDays||0);
  data.warrantyMonths=Number(data.warrantyMonths||0);
  return data;
}
async function post(url){
  statusEl.className="status";
  statusEl.textContent="Создаю документы…";
  const res=await fetch(url,{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify(payload())});
  if(!res.ok){
    const info=await res.json().catch(()=>({error:"Ошибка"}));
    throw new Error(info.error||"Не удалось выполнить операцию");
  }
  return res;
}
document.getElementById("download").addEventListener("click",async()=>{
  try{
    const res=await post("/api/generate");
    const blob=await res.blob();
    const a=document.createElement("a");
    a.href=URL.createObjectURL(blob);
    a.download="Документы.zip";
    a.click();
    URL.revokeObjectURL(a.href);
    statusEl.className="status ok";
    statusEl.textContent="Готово. ZIP-файл скачан.";
  }catch(e){
    statusEl.className="status err"; statusEl.textContent=e.message;
  }
});
document.getElementById("send").addEventListener("click",async()=>{
  try{
    if(!form.email.value) throw new Error("Введите e-mail для отправки.");
    const res=await post("/api/send");
    const info=await res.json();
    statusEl.className="status ok";
    statusEl.textContent=info.message||"Документы отправлены.";
  }catch(e){
    statusEl.className="status err"; statusEl.textContent=e.message;
  }
});
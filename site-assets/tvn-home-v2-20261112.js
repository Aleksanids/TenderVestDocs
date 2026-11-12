/* TenderVest site 22 — homepage interactions v2.  Client-side only, no external data sent. */
(function () {
 "use strict";
 function init() {
  var root=document.getElementById("tvn-home");
  if(!root || root.dataset.tvnReady==="1") return;
  root.dataset.tvnReady="1";
  document.documentElement.classList.add("tvn-js-ready");

  var scenarios={
   deadline:{
    beforeHead:"Сроки в разных местах",
    before:[["Файл «Закупки.xlsx»","неизвестно, актуален ли"],["Сообщение в чате","дедлайн обсуждали вчера"],["Напоминание вручную","нет общего контроля"]],
    afterHead:"Контрольная карточка",
    after:[["Срок подачи определен","виден ответственному"],["Задача назначена","с контрольной датой"],["Статус отслеживается","без поиска по чатам"]],
    problem:"Сроки зависят от переписки и памяти",outcome:"Сроки и ответственность видны",
    description:"Ответственный, дата подачи и состояние подготовки отображаются в одной карточке.",
    action:"Настроить контроль сроков",actionDescription:"Связать карточку тендера, задачи, сроки и ответственных в одном процессе.",
    link:"/services/tender-process-automation/"
   },
   documents:{
    beforeHead:"Документы разбросаны",
    before:[["Проект заявки","непонятна версия"],["Подтверждающие документы","у разных сотрудников"],["Контракт и акты","отдельная переписка"]],
    afterHead:"Комплект под контролем",
    after:[["Документы привязаны к закупке","версия и назначение"],["Проверка комплектности","ответственный назначен"],["Контракт и исполнение","в связанном процессе"]],
    problem:"Документы ищут перед дедлайном",outcome:"Документы доступны по закупке",
    description:"У каждой стадии свой комплект документов и понятный ответственный за проверку.",
    action:"Выстроить маршрут документов",actionDescription:"Связать заявку, подтверждения, согласования, контракт и закрывающие документы.",
    link:"/bitrix24/bitrix24-for-tenders/"
   },
   decision:{
    beforeHead:"Решение не зафиксировано",
    before:[["Конкуренция и цены","оценка на глаз"],["Требования заказчика","проверены частично"],["Причина участия","осталась в переписке"]],
    afterHead:"Обоснованное решение",
    after:[["Анализ условий","факты и ограничения"],["Экономика участия","согласованный ориентир"],["Участвовать / отказаться","записана причина"]],
    problem:"Непонятно, почему выбрали закупку",outcome:"Решение можно объяснить",
    description:"Требования, конкуренция и ограничения рассмотрены до участия; итог фиксируется.",
    action:"Настроить аналитический отбор",actionDescription:"Определить критерии оценки закупок, порядок согласования и нужные данные.",
    link:"/services/tender-analytics/"
   }
  };

  var range=root.querySelector("#tvn-range"),compare=root.querySelector("#tvn-compare-stage");
  var currentScenario="deadline";
  var status=root.querySelector("#tvn-compare-status");
  var outcome=root.querySelector("#tvn-compare-outcome");
  var desc=root.querySelector("#tvn-compare-desc");
  var actionTitle=root.querySelector("#tvn-compare-action");
  var actionDesc=root.querySelector("#tvn-compare-action-desc");
  var actionLink=root.querySelector("#tvn-compare-link");
  function text(id,value){var e=root.querySelector("#"+id);if(e)e.textContent=value;}
  function updateCompare(){
   if(!range||!compare)return;
   var v=Math.max(0,Math.min(100,Number(range.value)||0));
   compare.style.setProperty("--split",v+"%");
   range.setAttribute("aria-valuetext",v+" процентов состояния «После»");
   var scenario=scenarios[currentScenario];
   status.textContent=v<40?"До: разрозненная работа":v<65?"Переход: от ручного контроля к системе":"После: работа под контролем";
   outcome.textContent=v<40?scenario.problem:scenario.outcome;
   desc.textContent=v<40?"Данные, действия и сроки приходится проверять в разных местах.":scenario.description;
  }
  function chooseScenario(key){
   var data=scenarios[key];if(!data)return;
   currentScenario=key;
   root.querySelectorAll("[data-tvn-scenario]").forEach(function(b){b.setAttribute("aria-pressed",b.dataset.tvnScenario===key?"true":"false")});
   text("tvn-before-head",data.beforeHead);text("tvn-after-head",data.afterHead);
   for(var i=0;i<3;i++){
    text("tvn-before-"+(i+1),data.before[i][0]);text("tvn-before-meta"+(i+1),data.before[i][1]);
    text("tvn-after-"+(i+1),data.after[i][0]);text("tvn-after-meta"+(i+1),data.after[i][1]);
   }
   actionTitle.textContent=data.action;actionDesc.textContent=data.actionDescription;actionLink.href=data.link;
   updateCompare();
  }
  if(range){
   range.addEventListener("input",updateCompare);
   range.addEventListener("change",updateCompare);
  }
  root.querySelectorAll("[data-tvn-scenario]").forEach(function(button){
   button.addEventListener("click",function(){chooseScenario(button.dataset.tvnScenario)});
  });
  chooseScenario(currentScenario);

  var modules={
   search:{label:"Поиск и первичный отбор закупок",title:"Поиск закупок",desc:"Подходящие процедуры и очередь на первичную оценку"},
   analytics:{label:"Аналитика рынка и заказчиков",title:"Аналитика рынка",desc:"Заказчики, конкуренция и тенденции снижения"},
   card:{label:"Карточка закупки и решение об участии",title:"Карточка тендера",desc:"Условия, документы и зафиксированное решение"},
   deadlines:{label:"Роли, задачи и контроль сроков",title:"Роли и сроки",desc:"Ответственные, задачи и контрольные даты"},
   execution:{label:"Контракт, документы и исполнение",title:"Контракт и исполнение",desc:"Обязательства, приемка, закрывающие документы"},
   reporting:{label:"Отчетность для руководителя",title:"Управленческий отчет",desc:"Итоги участия, причины отказов и состояние процесса"}
  };
  var checkboxes=Array.from(root.querySelectorAll("[data-tvn-module]"));
  var preview=root.querySelector("#tvn-build-tiles");
  var count=root.querySelector("#tvn-selected-count");
  var buildLabel=root.querySelector("#tvn-build-label");
  var summary=root.querySelector("#tvn-build-summary");
  var mail=root.querySelector("#tvn-build-mail");
  function renderModules(){
   var chosen=checkboxes.filter(function(x){return x.checked}).map(function(x){return x.dataset.tvnModule});
   if(!preview)return;
   preview.replaceChildren();
   count.textContent=chosen.length+" из 6";
   if(!chosen.length){
    var empty=document.createElement("p");empty.className="tvn-build-empty";
    empty.textContent="Выберите хотя бы одну функцию, чтобы увидеть будущую структуру.";preview.appendChild(empty);
   }
   chosen.forEach(function(id){
    var m=modules[id];if(!m)return;
    var el=document.createElement("div");el.className="tvn-build-tile";
    var head=document.createElement("strong");head.textContent=m.title;
    var p=document.createElement("p");p.textContent=m.desc;
    el.append(head,p);preview.appendChild(el);
   });
   buildLabel.textContent=chosen.length===0?"Выбор не сделан":chosen.length<=2?"Стартовый контур":chosen.length<=4?"Расширенный контур":"Комплексный контур";
   summary.textContent=chosen.length ? chosen.map(function(id){return modules[id].title}).join(" · ") : "Отметьте процессы, которые нужно объединить.";
   var lines=chosen.map(function(id,i){return (i+1)+". "+modules[id].label}).join("\n");
   var body="Здравствуйте!\n\nХочу обсудить внедрение тендерного контура.\nВыбранные задачи:\n"+(lines||"Нужна помощь с выбором модулей.")+"\n\nКомпания:\nКонтакт для связи:\nКомментарий:\n";
   mail.href="mailto:a.rozyev@tender-vest.ru?subject="+encodeURIComponent("ТендерВест — оценка тендерного контура")+"&body="+encodeURIComponent(body);
   mail.setAttribute("aria-label","Открыть письмо с выбранными функциями: "+chosen.length);
  }
  checkboxes.forEach(function(el){el.addEventListener("change",renderModules)});
  renderModules();

  var stageList=[
   ["Поиск","Находим потенциально подходящие закупки и сохраняем исходные данные.","Список процедур для оценки"],
   ["Отбор","Проверяем требования, ограничения, условия и сроки.","Решение о допуске к дальнейшему анализу"],
   ["Решение","Сопоставляем условия, риски и экономическую целесообразность.","Обоснование: участвовать или отказаться"],
   ["Заявка","Собираем документы и проверяем комплектность.","Подготовленная и проверенная заявка"],
   ["Подача","Передаем заявку и подтверждаем регистрацию на площадке.","Подтвержденный факт подачи"],
   ["Торги","Отслеживаем итог процедуры и фиксируем цену.","Протоколы и результат участия"],
   ["Контракт","Проверяем обязательства и сроки подписания.","Подписанный договор и план исполнения"],
   ["Исполнение","Контролируем обязательства, акты, УПД и оплату.","Подтвержденное исполнение и расчеты"]
  ];
  function setStage(index){
   var idx=Math.max(0,Math.min(stageList.length-1,index));
   var data=stageList[idx];
   root.querySelectorAll("[data-tvn-stage]").forEach(function(b){b.setAttribute("aria-pressed",Number(b.dataset.tvnStage)===idx?"true":"false")});
   text("tvn-stage-index",String(idx+1).padStart(2,"0"));
   text("tvn-stage-title",data[0]);text("tvn-stage-info",data[1]);text("tvn-stage-result",data[2]);
  }
  root.querySelectorAll("[data-tvn-stage]").forEach(function(b){b.addEventListener("click",function(){setStage(Number(b.dataset.tvnStage))})});
  setStage(0);

  if(!window.matchMedia("(prefers-reduced-motion: reduce)").matches &&
      typeof IntersectionObserver!=="undefined"){
    root.classList.add("tvn-motion");
    var observer=new IntersectionObserver(function(entries){
     entries.forEach(function(entry){if(entry.isIntersecting){entry.target.classList.add("is-visible");observer.unobserve(entry.target)}});
    },{threshold:.08});
    root.querySelectorAll(".tvn-reveal").forEach(function(el){observer.observe(el)});
  }
  var help=root.querySelector("#tvn-help-widget");
  if(help){document.addEventListener("keydown",function(event){if(event.key==="Escape")help.open=false})}
 }
 if(document.readyState==="loading"){document.addEventListener("DOMContentLoaded",init,{once:true})}
 else{init()}
})();
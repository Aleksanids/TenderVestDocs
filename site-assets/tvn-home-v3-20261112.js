/* TenderVest homepage v3: interactive split + tender workflow configurator. No external tracking or client data. */
(function(){
 "use strict";
 var ROOT_ID="tvn-home";
 function start(){
  var root=document.getElementById(ROOT_ID);
  if(!root || root.dataset.tvnV3Ready==="1")return;
  root.dataset.tvnV3Ready="1";
  document.documentElement.classList.add("tvn-js-ready");
  function find(s){return root.querySelector(s)}
  function setText(id,val){var e=find("#"+id);if(e)e.textContent=val}
  var scenarios={
   deadline:{
    beforeHead:"Сроки в разных местах",
    before:[["Тендеры_финал.xlsx","неясно, актуальна ли версия"],["Кто отслеживает подачу?","ответственный не назначен"],["Не пропустить дедлайн","личное напоминание"]],
    afterHead:"Сроки на контроле",
    after:[["Карточка закупки","извещение и дата"],["Ответственный назначен","задача и срок"],["Действия отслеживаются","история и статус"]],
    problem:"Дедлайны в таблицах и переписке",
    outcome:"У каждого срока есть ответственный",
    explanation:"Дата подачи, задача и следующий шаг связаны в одной карточке.",
    action:"Настроить контроль сроков",
    how:"Согласовать правила отбора, ответственных и контрольные точки.",
    url:"/services/tender-process-automation/"
   },
   documents:{
    beforeHead:"Документы разбросаны",
    before:[["Заявка_финал_3.docx","версия под вопросом"],["Где выписка и гарантия?","файлы в разных чатах"],["Акт после поставки","искать отдельно"]],
    afterHead:"Документы связаны с закупкой",
    after:[["Заявка и комплект","версия и статус"],["Согласование","назначен проверяющий"],["Контракт, акты, УПД","в общем маршруте"]],
    problem:"Документы приходится искать",
    outcome:"Комплект документов под контролем",
    explanation:"У каждой стадии есть документ, ответственный и понятный статус.",
    action:"Настроить маршрут документов",
    how:"Связать заявку, проверки, согласования и исполнение.",
    url:"/bitrix24/bitrix24-for-tenders/"
   },
   decision:{
    beforeHead:"Решение без доказательств",
    before:[["Снижение «примерно»","без анализа участников"],["Требования заказчика","проверены не все"],["Почему участвуем?","не зафиксировано"]],
    afterHead:"Решение с обоснованием",
    after:[["Условия закупки","проверены ограничения"],["Конкуренция","изучена история"],["Участвовать / отказаться","причина сохранена"]],
    problem:"Неясно, почему выбрали закупку",
    outcome:"Решение можно проверить",
    explanation:"Условия, конкуренция и риски сопоставлены до подачи.",
    action:"Ввести аналитический отбор",
    how:"Согласовать критерии оценки и порядок решения об участии.",
    url:"/services/tender-analytics/"
   }
  };
  var compare=find("#tvn-compare-stage"),range=find("#tvn-range");
  var chosenScenario="deadline";
  function renderSplit(){
   if(!range||!compare)return;
   var n=Math.min(100,Math.max(0,Number(range.value)||0));
   compare.style.setProperty("--split",n+"%");
   range.setAttribute("aria-valuetext","Граница в "+n+" процентах: слева до, справа после");
   var d=scenarios[chosenScenario];
   setText("tvn-compare-status",n<40?"ПОСЛЕ / БОЛЬШЕ СИСТЕМЫ":n<=60?"СРАВНИТЕ ДВА СОСТОЯНИЯ":"ДО / БОЛЬШЕ РАЗРОЗНЕННОСТИ");
   setText("tvn-compare-outcome",n>60?d.problem:d.outcome);
   setText("tvn-compare-desc",n>60?"Данные и действия находятся в отдельных источниках, решение приходится восстанавливать вручную.":d.explanation);
  }
  function scenario(key){
   var d=scenarios[key];if(!d)return;chosenScenario=key;
   root.querySelectorAll("[data-tvn-scenario]").forEach(function(e){e.setAttribute("aria-pressed",e.dataset.tvnScenario===key?"true":"false")});
   setText("tvn-before-head",d.beforeHead);setText("tvn-after-head",d.afterHead);
   for(var i=0;i<3;i++){
    var j=i+1;
    setText("tvn-before-"+j,d.before[i][0]);setText("tvn-before-meta"+j,d.before[i][1]);
    setText("tvn-after-"+j,d.after[i][0]);setText("tvn-after-meta"+j,d.after[i][1]);
   }
   setText("tvn-compare-action",d.action);
   setText("tvn-compare-action-desc",d.how);
   var link=find("#tvn-compare-link");if(link)link.href=d.url;
   renderSplit();
  }
  if(range){range.addEventListener("input",renderSplit);range.addEventListener("change",renderSplit)}
  root.querySelectorAll("[data-tvn-scenario]").forEach(function(b){b.addEventListener("click",function(){scenario(b.dataset.tvnScenario)})});
  scenario("deadline");

  var moduleData={
   search:{label:"Поиск и отбор закупок",title:"Новая закупка",hint:"Ожидает первичной оценки",status:"● Извещение сохранено",lane:"search"},
   analytics:{label:"Аналитика рынка",title:"Карта конкурентной среды",hint:"Заказчики · участники · снижение",status:"● Данные для решения",lane:"search"},
   card:{label:"Карточка тендера",title:"Решение об участии",hint:"Требования и ограничения",status:"● На согласовании",lane:"decision"},
   deadlines:{label:"Роли и контроль сроков",title:"Контроль срока подачи",hint:"Ответственный и дата",status:"● Задача назначена",lane:"decision"},
   execution:{label:"Контракт и исполнение",title:"Контракт к исполнению",hint:"Акты, УПД, оплата",status:"● Контроль обязательств",lane:"contract"},
   reporting:{label:"Отчетность руководителя",title:"Панель результата",hint:"Выбор, участие, исполнение",status:"● Управленческий отчет",lane:"contract"}
  };
  var boxes=Array.from(root.querySelectorAll("[data-tvn-module]"));
  var lanes=["search","decision","contract"];
  var c=find("#tvn-selected-count"),dashboardCount=find("#tvn3-dashboard-count"),dashboardCtrl=find("#tvn3-dashboard-control"),dashboardNext=find("#tvn3-dashboard-next");
  var summary=find("#tvn-build-summary"),label=find("#tvn-build-label"),mail=find("#tvn-build-mail");
  function make(tag,cls,content){var e=document.createElement(tag);if(cls)e.className=cls;if(content)e.textContent=content;return e}
  function renderWorkflow(){
   var selected=boxes.filter(function(e){return e.checked}).map(function(e){return e.dataset.tvnModule});
   if(c)c.textContent=selected.length+" из 6";
   if(dashboardCount)dashboardCount.textContent=String(selected.length).padStart(2,"0")+" / 06";
   var activeLanes=[];
   lanes.forEach(function(lane){
    var container=find("#tvn3-lane-"+lane),count=find("#tvn3-lane-count-"+lane);
    if(!container)return;
    var values=selected.filter(function(k){return moduleData[k]&&moduleData[k].lane===lane});
    if(count)count.textContent=String(values.length).padStart(2,"0");
    if(values.length)activeLanes.push(lane==="search"?"Поиск":lane==="decision"?"Решение":"Исполнение");
    container.replaceChildren();
    if(!values.length){
     container.appendChild(make("div","tvn3-workspace-empty",lane==="search"?"Включите поиск или аналитику":lane==="decision"?"Включите карточку или сроки":"Включите исполнение или отчетность"));
    }else{
     values.forEach(function(k){
      var d=moduleData[k],card=make("article","tvn3-item");
      card.appendChild(make("strong","",d.title));
      card.appendChild(make("span","",d.hint));
      card.appendChild(make("small","",d.status));
      container.appendChild(card);
     });
    }
   });
   if(dashboardCtrl)dashboardCtrl.textContent=activeLanes.join(" / ")||"Не выбрано";
   var next=selected.includes("execution")?"Проверить исполнение":selected.includes("deadlines")?"Отследить срок":selected.includes("card")?"Согласовать решение":selected.includes("analytics")?"Сравнить участников":selected.includes("search")?"Отобрать закупку":selected.includes("reporting")?"Смотреть показатели":"Выбрать процесс";
   if(dashboardNext)dashboardNext.textContent=next;
   if(label)label.textContent=selected.length===0?"Выбор не сделан":selected.length<=2?"Стартовый контур":selected.length<=4?"Контур с автоматизацией":"Комплексный контур";
   if(summary)summary.textContent=selected.length?selected.map(function(k){return moduleData[k].label}).join(" · "):"Выберите задачи, которые хотите объединить.";
   if(mail){
    var lines=selected.map(function(k,i){return(i+1)+". "+moduleData[k].label}).join("\n");
    var body="Здравствуйте!\n\nХочу обсудить организацию тендерной работы. Выбрал(а) следующие модули:\n"+(lines||"Требуется консультация по составу решения")+"\n\nСфера деятельности компании:\nЧто сейчас не работает:\nЖелаемый результат:\nКонтакт для связи:\n";
    mail.href="mailto:a.rozyev@tender-vest.ru?subject="+encodeURIComponent("ТендерВест — состав тендерного контура")+"&body="+encodeURIComponent(body);
   }
  }
  boxes.forEach(function(b){b.addEventListener("change",renderWorkflow)});
  renderWorkflow();
  var widget=find("#tvn-help-widget");if(widget){document.addEventListener("keydown",function(e){if(e.key==="Escape")widget.open=false})}
  if(!window.matchMedia("(prefers-reduced-motion: reduce)").matches && typeof IntersectionObserver!=="undefined"){
   root.classList.add("tvn-motion");
   var ob=new IntersectionObserver(function(entries){
    entries.forEach(function(ent){if(ent.isIntersecting){ent.target.classList.add("is-visible");ob.unobserve(ent.target)}})
   },{threshold:.07});
   root.querySelectorAll(".tvn-reveal").forEach(function(el){ob.observe(el)});
  }
 }
 if(document.readyState==="loading"){document.addEventListener("DOMContentLoaded",start,{once:true})}
 else{start()}
})();
/*Код для вибору блоків та активності і не активності кнопки */
$(document).ready(function(){
    $(".block_2 div").click(function(){
      $(this).parent().find('div').removeClass("selected");
      $(this).toggleClass("selected");
      checkForm();
    });
  
    $("._stydent").on('input', function() {
      checkForm();
    });
  
    function checkForm() {
      var allFilled = true;
      $("._stydent").each(function() {
        if($(this).val() == '') {
          allFilled = false;
        }
      });
  
      $(".block_2").each(function() {
        if($(this).find(".selected").length == 0) {
          allFilled = false;
        }
      });
  
      if(allFilled) {
        $(".button_stydent_form a").addClass("active");
      } else {
        $(".button_stydent_form a").removeClass("active");
      }
    }
  });

/*плавна прокрутка */
var links = document.querySelectorAll('a[href^="#"]');
  
links.forEach(function(link) {
  link.addEventListener('click', function(e) {
    e.preventDefault();

    var targetId = this.getAttribute('href').substring(1);
    var targetElement = document.getElementById(targetId);

    window.scrollTo({
      top: targetElement.offsetTop,
      behavior: 'smooth'
    });
  });
});

/* Анімація при скролі*/
const animItems = document.querySelectorAll('.anim');

if (animItems.length > 0) {
    window.addEventListener('scroll', animOnScroll);

    function animOnScroll() {
        for (let index = 0; index < animItems.length; index++) {
            const animItem = animItems[index];
            const animItemHeight = animItem.offsetHeight;
            const animItemOffset = offset(animItem).top;
            const animStart = 4;

            let animItemPoint = window.innerHeight - animItemHeight / animStart;
            if (animItemHeight > window.innerHeight) {
                animItemPoint = window.innerHeight - window.innerHeight / animStart;
            }

            if ((window.pageYOffset > animItemOffset - animItemPoint) && window.pageYOffset < (animItemOffset + animItemHeight)) {
                animItem.classList.add('_active');
            } else {
                animItem.classList.remove('_active');
            }
        }
    }

    function offset(el) {
        const rect = el.getBoundingClientRect(),
            scrollLeft = window.pageXOffset || document.documentElement.scrollLeft,
            scrollTop = window.pageYOffset || document.documentElement.scrollTop;
        return { top: rect.top + scrollTop, left: rect.left + scrollLeft };
    }

    animOnScroll(); // Викликаємо функцію один раз після завантаження сторінки, щоб встановити класи анімації.
}


  








/*структура OK і ВК */
/*Спитати:
  1) За кількість годин Всього та Аудиторні години
  2)Лекції чи просто задати чи якось рахуються
  3) Чи об'єднувати лабораторні практичні
*/
let tmp={
  "conVK":"ОК1",
  "name":"Економічна теорія",
  "countCredit":3,
  "course":2,
  "semester":1,
  "hoursInWeek":3,
  "countWeek":16,
  "formOfControl":"залік",
  "lectures":35,
  "practicalLaboratory":0,
  "seminar":0

};
const P={
  "conOK":"ОК34",
  "name":"Навчальна практика з інформаційних технологій",
  "countCredit":2,
  "course":2,
  "semester":2,
  "lenght":2
};


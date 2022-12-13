try {
    var j = document.createElement("script");
    j.type = "text/javascript";
    j.src = "//odnaknopka.ru/ok9.js";
    if(document.querySelector("body")){
    document.querySelector("body").appendChild(j);
    } else if (document.querySelector("head")) {
    document.querySelector("head").appendChild(j);
    }
    } catch(e) {
    }
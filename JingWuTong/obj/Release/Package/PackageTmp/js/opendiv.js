
function lqopen_newDrownFrame(id, father, frameSrc,width,height, headText, src, headBackUrl, closeEventName) {
    if (typeof (father) == "string") {
        father = document.getElementById(father);
    }
    if (!src) {
        src = "MyCommonJS/Source/openDiv/infowindow_close.gif";
    }
    if (!headBackUrl) {
        headBackUrl = "Images/bgbj.gif";
    }
    if (closeEventName) {
        closeEventName = "window.parent." + closeEventName + "();";
    }
    else {
        closeEventName = "";
    }
    if (!document.getElementById(id)) {
        var frame = document.createElement("iframe");
        frame.id = id + "_ifr";
        frame.src = frameSrc;
        frame.width = "100%";
        frame.height = "100%";
        frame.allowTransparency = "true";
        frame.frameBorder = "no";
        if (frame.id != "pcstatusinfo_ifr") {
         //   frame.style.paddingBottom = "41px";
            frame.scrolling = "no";
        }
        else {
            frame.scrolling = "no";
        }
        var div = document.createElement("div");
        div.id = id;
        div.style.position = "absolute";
        div.style.overflow = "hidden";
        div.style.display = "block";
        div.style.top = "100";
        div.style.left = "100";
        div.style.width = width;
        div.style.height = height;
        if (closeEventName)
        { }
        div.appendChild(frame);
        father.appendChild(div);
    }
}
//ajax调用防止出现返回缓存数据
function TimeGet() {
    var myDate = new Date();
    return "_hours:" + myDate.getHours() + "_minutes:" + myDate.getMinutes() + "_seconds:" + myDate.getSeconds() + "_milliseconds:" + myDate.getMilliseconds();
}


function mycallfunction(calltype, width, height, lq_id, zindex, newwindows) {

    var div = document.getElementById(calltype);
    var ifr = document.getElementById(calltype + "_ifr");
    if (!newwindows) {
        var hrsrc = "http://10.8.52.241:8080/lqnew/opePages/" + calltype + ".aspx?time=1";
        //var hrsrc = "http://localhost:62758/lqnew/opePages/" + calltype + ".aspx?time=1";
        
    }
    else {
        var hrsrc = "http://10.8.52.241:8080/lqnew/opePages/" + calltype + ".aspx?time=" + TimeGet();
        //var hrsrc = "http://localhost:62758/lqnew/opePages/" + calltype + ".aspx?time=" + TimeGet();
    }
    if (!div) {

        if (!lq_id) { lqopen_newDrownFrame(calltype, document.body, hrsrc, width, height); }
        else { lqopen_newDrownFrame(calltype, document.body, hrsrc + "&id=" + lq_id, width, height); }

    }
    else {
        ifr.src = "about:blank"; removeChildSafe(div); ifr = null; div = null;
    }


}


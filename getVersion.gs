function getVersion() {
  var soft = "aba WBS manager Tool.";
  var version = "Version: 0.0.1";
  var cpright = "(c) 2016 aba Co.,Ltd.";
  var contact = "";
  
  var LF ="\\n";  //String.fromCharCode(10);
  var LF2 = LF + LF;
  
  var msg = soft + LF2 + version + LF + cpright + LF2 + contact;
  Browser.msgBox(msg, Browser.Buttons.OK);
}

function run() {
    var app = Application.currentApplication();
    app.includeStandardAdditions = true;
    var content = app.doShellScript("cat '/Users/swdjsj/.gemini/antigravity-ide/brain/3c8c569b-2783-45f9-9363-0d5f1c6454a2/.system_generated/tasks/task-288.log'");
    var data = JSON.parse(content);
    
    var seonjunData = data.seonjun;
    if (seonjunData) {
        for (var i = 0; i < seonjunData.length; i++) {
            var item = seonjunData[i];
            var str = JSON.stringify(item);
            if (str.indexOf('아이폰') !== -1 || str.indexOf('피부과') !== -1 || str.indexOf('제습기') !== -1) {
                console.log(str);
            }
        }
    } else {
        console.log("No seonjun property found");
    }
}

#target illustrator

// 确保有活动文档
if (app.documents.length > 0) {
    var doc = app.activeDocument;

    // 定义offset变量
    var offset = 5; // 你可以根据需要调整这个值

    // 定义新的填充颜色
    var newColor = new RGBColor();
    newColor.red = 255;
    newColor.green = 255;
    newColor.blue = 255;

    // 应用Outline效果的函数
    function applyOutlineEffect(path)
    {
        var xmlstring = '<LiveEffect name="Adobe Offset Path"><Dict data="R mlim 4 R ofst ' + offset + ' I jntp 2 "/></LiveEffect>';
        path.applyEffect(xmlstring);
        path.opacity = 100;  
        path.fillColor = newColor;
    }

    // 应用投影效果的函数
    function applyDropShadow(path)
    {
        var shadowXML = '<LiveEffect name="Adobe Drop Shadow">' +
                        '<Dict data="B pair 1 R opac 1 R dark 20 R horz 2 R blur 3 I csrc 1 I blnd 0 R vert 2 B usePSLBlur 1 I Adobe Effect Expand Before Version 16 ">' +
                        '<Entry name="sclr" valueType="F">' +
                        '<Fill darkness="0,0,0"/>' +
                        '</Entry>' +
                        '</Dict>' +
                        '</LiveEffect>';
        path.applyEffect(shadowXML);
    }

    // 处理单个剪切组的函数
    function processClippingGroup(clippingGroup) {
        if (clippingGroup.pathItems.length > 0) {
            // 复制第一个路径并移出剪切组
            var copiedPath = clippingGroup.pathItems[0].duplicate();
            copiedPath.moveToBeginning(doc);
            
            // 使用zOrder将复制的路径下移
            copiedPath.zOrder(ZOrderMethod.SENDBACKWARD);
            
            // 应用Outline效果
            applyOutlineEffect(copiedPath);
            
            // 应用投影效果
            applyDropShadow(copiedPath);
            
            // 命名新路径
            copiedPath.name = "扩展边框";
            
            // 创建新组并将原剪切组和新路径添加到其中
            var newGroup = doc.groupItems.add();
            clippingGroup.move(newGroup, ElementPlacement.PLACEATBEGINNING);
            copiedPath.move(newGroup, ElementPlacement.PLACEATEND);
            
            return newGroup;
        }
        return null;
    }

    // 获取所有选中的对象
    var selection = doc.selection;
    
    if (selection.length > 0) {
        // 创建一个数组来存储新创建的组
        var newGroups = [];
        
        // 使用 try-catch 块来确保脚本不会因为单个错误而完全停止
        for (var i = 0; i < selection.length; i++) {
            try {
                var item = selection[i];
                if (item.typename === "GroupItem" && item.clipped) {
                    var newGroup = processClippingGroup(item);
                    if (newGroup) {
                        newGroups.push(newGroup);
                    }
                }
            } catch (e) {
                alert("处理第 " + (i + 1) + " 个对象时出错: " + e.message);
            }
        }
        
        // 选中所有新创建的组
        doc.selection = newGroups;
        
        //alert("处理完成！共处理 " + newGroups.length + " 个剪切组。");
    } else {
        alert("请先选择至少一个剪切组！");
    }
} else {
    alert("请先打开一个文档！");
}

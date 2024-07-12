#target illustrator

// 调试模式开关
var DEBUG_MODE = false;

// 用于收集调试信息的全局变量
var debugInfo = "";

// 调试信息函数
function debug(message) {
    if (DEBUG_MODE) {
        debugInfo += message + "\n";
    }
}

// 显示调试信息的函数
function showDebugInfo() {
    if (DEBUG_MODE && debugInfo !== "") {
        alert(debugInfo);
        debugInfo = ""; // 清空调试信息
    }
}

// 主函数
function main() {
    if (app.documents.length == 0) {
        alert("请先打开一个文档！");
        return;
    }

    var doc = app.activeDocument;
    var selection = doc.selection;

    if (selection.length == 0) {
        alert("请先选择一些对象！");
        return;
    }

    debug("开始处理选中的对象");

    // 1. 检查可以执行剪切蒙版的对象
    var validObjects = [];
    for (var i = 0; i < selection.length; i++) {
        var item = selection[i];
        if (item.typename == "PathItem" || item.typename == "CompoundPathItem" || isClippingGroup(item)) {
            validObjects.push(item);
        }
    }

    debug("有效对象数量: " + validObjects.length);

    if (validObjects.length == 0) {
        alert("没有找到可以执行剪切蒙版的对象！");
        return;
    }

    // 限制处理的对象数量
    var maxObjects = 50;
    if (validObjects.length > maxObjects) {
        alert("选择的对象过多，将只处理前 " + maxObjects + " 个对象。");
        validObjects = validObjects.slice(0, maxObjects);
    }

    // 2. 计算对象的长宽比例和类型
    var objectInfo = [];
    for (var i = 0; i < validObjects.length; i++) {
        var obj = validObjects[i];
        var bounds = getPathBounds(obj);
        var width = bounds[2] - bounds[0];
        var height = bounds[1] - bounds[3];
        var ratio = width / height;
        var type = getShapeType(ratio);
        objectInfo.push({ratio: ratio, type: type, index: i});
    }

    var objectInfoStr = "";
    for (var i = 0; i < objectInfo.length; i++) {
        objectInfoStr += "类型: " + objectInfo[i].type + ", 比例: " + objectInfo[i].ratio.toFixed(2);
        if (i < objectInfo.length - 1) objectInfoStr += "; ";
    }
    debug("对象信息: " + objectInfoStr);

    // 3. 弹出对话框选择图片
    var images = File.openDialog("选择图片", "*.png;*.jpg;*.jpeg", true);

    if (images == null || images.length == 0) {
        alert("没有选择图片！");
        return;
    }

    debug("选择的图片数量: " + images.length);

    // 4. 计算图片的长宽比例和类型
    var imageInfo = [];
    for (var i = 0; i < images.length; i++) {
        var ratio = getImageRatio(images[i]);
        var type = getShapeType(ratio);
        imageInfo.push({ratio: ratio, type: type, index: i});
    }

    var imageInfoStr = "";
    for (var i = 0; i < imageInfo.length; i++) {
        imageInfoStr += "类型: " + imageInfo[i].type + ", 比例: " + imageInfo[i].ratio.toFixed(2);
        if (i < imageInfo.length - 1) imageInfoStr += "; ";
    }
    debug("图片信息: " + imageInfoStr);

    // 5. 使用改进的匹配算法找到最优匹配
    var matching = improvedMatching(objectInfo, imageInfo);

    debug("匹配结果: " + matching.join(", "));

    // 6. 根据匹配结果创建剪切蒙版
    var progressBar = createProgressBar(matching.length);
    for (var i = 0; i < matching.length; i++) {
        if (matching[i] !== -1) {
            debug("正在处理对象 " + i + " 和图片 " + matching[i]);
            placeImageInPath(doc, validObjects[i], images[matching[i]]);
        } else {
            debug("对象 " + i + " 没有匹配的图片");
        }
        updateProgressBar(progressBar, i + 1);
    }
    progressBar.close();

    showDebugInfo(); // 显示收集的调试信息
    //alert("操作完成！");
}

// 获取形状类型
function getShapeType(ratio) {
    var tolerance = 0.1; // 允许的误差范围
    if (Math.abs(ratio - 1) <= tolerance) {
        return "square";
    } else if (ratio > 1) {
        return "landscape";
    } else {
        return "portrait";
    }
}

// 改进的匹配算法
function improvedMatching(objectInfo, imageInfo) {
    var matching = [];
    var usedImages = {};

    // 首先匹配方形
    matchShapes(objectInfo, imageInfo, usedImages, matching, "square");

    // 然后匹配横向矩形
    matchShapes(objectInfo, imageInfo, usedImages, matching, "landscape");

    // 最后匹配纵向矩形
    matchShapes(objectInfo, imageInfo, usedImages, matching, "portrait");

    // 处理剩余未匹配的对象
    for (var i = 0; i < objectInfo.length; i++) {
        if (matching[objectInfo[i].index] === undefined) {
            var bestMatch = findBestUnusedMatch(objectInfo[i], imageInfo, usedImages);
            if (bestMatch !== -1) {
                matching[objectInfo[i].index] = bestMatch;
                usedImages[bestMatch] = true;
            } else {
                matching[objectInfo[i].index] = -1;
            }
        }
    }

    return matching;
}

// 匹配特定形状
function matchShapes(objectInfo, imageInfo, usedImages, matching, shapeType) {
    var objects = [];
    for (var i = 0; i < objectInfo.length; i++) {
        if (objectInfo[i].type === shapeType) {
            objects.push(objectInfo[i]);
        }
    }
    var images = [];
    for (var i = 0; i < imageInfo.length; i++) {
        if (imageInfo[i].type === shapeType && !usedImages[imageInfo[i].index]) {
            images.push(imageInfo[i]);
        }
    }

    for (var i = 0; i < objects.length; i++) {
        var bestMatch = findBestMatch(objects[i], images, usedImages);
        if (bestMatch !== -1) {
            matching[objects[i].index] = bestMatch;
            usedImages[bestMatch] = true;
        }
    }
}

// 找到最佳匹配
function findBestMatch(object, images, usedImages) {
    var bestMatch = -1;
    var minDiff = Infinity;

    for (var i = 0; i < images.length; i++) {
        if (!usedImages[images[i].index]) {
            var diff = Math.abs(object.ratio - images[i].ratio);
            if (diff < minDiff) {
                minDiff = diff;
                bestMatch = images[i].index;
            }
        }
    }

    return bestMatch;
}

// 为未匹配的对象找到最佳未使用的图片
function findBestUnusedMatch(object, imageInfo, usedImages) {
    var bestMatch = -1;
    var minDiff = Infinity;

    for (var i = 0; i < imageInfo.length; i++) {
        if (!usedImages[i]) {
            var diff = Math.abs(object.ratio - imageInfo[i].ratio);
            if (diff < minDiff) {
                minDiff = diff;
                bestMatch = i;
            }
        }
    }

    return bestMatch;
}

// 检查是否为剪切组
function isClippingGroup(item) {
    return item.typename == "GroupItem" && item.clipped;
}

// 查找剪切路径
function findClippingPath(group) {
    for (var i = 0; i < group.pageItems.length; i++) {
        if (group.pageItems[i].clipping) {
            return group.pageItems[i];
        }
    }
    return null;
}

// 获取路径边界
function getPathBounds(pathItem) {
    if (isClippingGroup(pathItem)) {
        var clippingPath = findClippingPath(pathItem);
        return clippingPath ? clippingPath.geometricBounds : pathItem.geometricBounds;
    }
    return pathItem.geometricBounds;
}

// 获取图片比例
function getImageRatio(file) {
    var image = app.activeDocument.placedItems.add();
    image.file = new File(decodeURI(file.absoluteURI));
    var ratio = image.width / image.height;
    image.remove();
    return ratio;
}

// 将图片放置到路径中并创建剪切蒙版
function placeImageInPath(doc, pathItem, file) {
    var clippingPath;
    var clippingGroup = null;

    // 如果pathItem是剪切组，找到剪切路径并移除内部所有项目
    if (isClippingGroup(pathItem)) {
        clippingGroup = pathItem;
        clippingPath = findClippingPath(clippingGroup);

        // 移除剪切组内的所有非剪切路径项目
        for (var i = clippingGroup.pageItems.length - 1; i >= 0; i--) {
            if (!clippingGroup.pageItems[i].clipping) {
                clippingGroup.pageItems[i].remove();
            }
        }
    } else {
        clippingPath = pathItem;
    }

    // 将新图片文件放置到文档中
    var placedItem = doc.placedItems.add();
    placedItem.file = new File(decodeURI(file.absoluteURI));

    // 调整放置的图片大小以适应路径项目大小
    var pathBounds = getPathBounds(clippingPath);
    var pathWidth = pathBounds[2] - pathBounds[0];
    var pathHeight = pathBounds[1] - pathBounds[3];

    var imageBounds = placedItem.geometricBounds;
    var imageWidth = imageBounds[2] - imageBounds[0];
    var imageHeight = imageBounds[1] - imageBounds[3];

    // 计算缩放比例以使图片的较短边适应路径
    var widthRatio = pathWidth / imageWidth;
    var heightRatio = pathHeight / imageHeight;
    var scaleRatio = Math.max(widthRatio, heightRatio); // 使用max确保较短边适应

    // 使用正确的缩放因子调整放置项目的大小
    placedItem.resize(scaleRatio * 100, scaleRatio * 100, true, true, true, true, scaleRatio * 100, Transformation.DOCUMENTORIGIN);

    // 将图片居中放置在路径内
    var newImageBounds = placedItem.geometricBounds;
    var newImageWidth = newImageBounds[2] - newImageBounds[0];
    var newImageHeight = newImageBounds[1] - newImageBounds[3];

    var deltaX = (pathWidth - newImageWidth) / 2;
    var deltaY = (pathHeight - newImageHeight) / 2;

    placedItem.left = pathBounds[0] + deltaX;
    placedItem.top = pathBounds[1] - deltaY;

    // 特殊处理复合路径
    if (pathItem.typename == "CompoundPathItem") {
        // 选择复合路径和图片
        doc.selection = null;
        pathItem.selected = true;
        placedItem.selected = true;
        
        // 执行"剪切蒙版"命令
        app.executeMenuCommand('makeMask');
    } else {
        // 对于普通路径和现有的剪切组，使用原来的方法
        clippingPath.zOrder(ZOrderMethod.BRINGTOFRONT);

        if (clippingGroup) {
            placedItem.move(clippingGroup, ElementPlacement.PLACEATEND);
        } else {
            clippingGroup = doc.groupItems.add();
            placedItem.move(clippingGroup, ElementPlacement.PLACEATEND);
            clippingPath.move(clippingGroup, ElementPlacement.PLACEATBEGINNING);
            clippingGroup.clipped = true;
        }
    }
}

// 创建进度条
function createProgressBar(total) {
    var progressBar = new Window("palette", "处理进度");
    progressBar.progressBar = progressBar.add("progressbar", undefined, 0, total);
    progressBar.progressBar.preferredSize.width = 300;
    progressBar.show();
    return progressBar;
}

// 更新进度条
function updateProgressBar(progressBar, value) {
    progressBar.progressBar.value = value;
    progressBar.update();
}

// 运行脚本
main();

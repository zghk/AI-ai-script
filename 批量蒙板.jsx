#target illustrator

var doc = app.activeDocument;
var sel = doc.selection;

if (sel.length == 0) {
    alert("请选择至少一个对象！");
} else {
    main();
}

function main() {
    var validItems = getValidItems(sel);
    if (validItems.length == 0) {
        alert("没有找到可以创建剪切蒙版的对象！");
        return;
    }

    var images = File.openDialog("选择图片", "*.jpg;*.jpeg;*.png;*.gif", true);
    if (images == null || images.length == 0) return;

    var itemRatios = calculateItemRatios(validItems);
    var imageRatios = calculateImageRatios(images);

    // 创建进度条
    var progressBar = new ProgressBar("处理进度", validItems.length);

    matchAndCreateMasks(validItems, images, itemRatios, imageRatios, progressBar);

    // 关闭进度条
    progressBar.close();

    alert("操作完成！");
}

function getValidItems(selection) {
    var validItems = [];
    for (var i = 0; i < selection.length; i++) {
        var item = selection[i];
        if (item.typename == "PathItem" || item.typename == "CompoundPathItem" || 
            (item.typename == "GroupItem" && item.clipped)) {
            validItems.push(item);
        }
    }
    return validItems;
}

function calculateItemRatios(items) {
    var ratios = [];
    for (var i = 0; i < items.length; i++) {
        var item = items[i];
        var bounds = item.geometricBounds;
        ratios.push((bounds[2] - bounds[0]) / (bounds[3] - bounds[1]));
    }
    return ratios;
}

function calculateImageRatios(images) {
    var ratios = [];
    for (var i = 0; i < images.length; i++) {
        var image = images[i];
        var tempItem = doc.placedItems.add();
        tempItem.file = image;
        ratios.push(tempItem.width / tempItem.height);
        tempItem.remove();
    }
    return ratios;
}

function matchAndCreateMasks(items, images, itemRatios, imageRatios, progressBar) {
    var usedImages = [];
    var itemIndices = [];
    for (var i = 0; i < items.length; i++) {
        itemIndices.push(i);
    }

    var processedCount = 0;
    while (itemIndices.length > 0 && images.length > 0) {
        var bestMatch = findBestMatch(itemRatios, imageRatios);
        if (bestMatch.itemIndex !== -1 && bestMatch.imageIndex !== -1) {
            createClippingMask(items[itemIndices[bestMatch.itemIndex]], images[bestMatch.imageIndex]);
            usedImages.push(images[bestMatch.imageIndex]);
            itemIndices.splice(bestMatch.itemIndex, 1);
            itemRatios.splice(bestMatch.itemIndex, 1);
            images.splice(bestMatch.imageIndex, 1);
            imageRatios.splice(bestMatch.imageIndex, 1);

            // 更新进度条
            processedCount++;
            progressBar.update(processedCount);
        } else {
            break;
        }
    }
    return usedImages;
}

function findBestMatch(itemRatios, imageRatios) {
    var bestItemIndex = -1;
    var bestImageIndex = -1;
    var minDiff = Infinity;

    for (var i = 0; i < itemRatios.length; i++) {
        for (var j = 0; j < imageRatios.length; j++) {
            var diff = Math.abs(itemRatios[i] - imageRatios[j]);
            if (diff < minDiff) {
                minDiff = diff;
                bestItemIndex = i;
                bestImageIndex = j;
            }
        }
    }

    return { itemIndex: bestItemIndex, imageIndex: bestImageIndex };
}

function createClippingMask(item, image) {
    var placedItem = doc.placedItems.add();
    placedItem.file = image;

    // 获取路径的边界
    var pathBounds = item.geometricBounds;
    var pathWidth = pathBounds[2] - pathBounds[0];
    var pathHeight = pathBounds[3] - pathBounds[1];

    // 获取图片的边界
    var imageBounds = placedItem.geometricBounds;
    var imageWidth = imageBounds[2] - imageBounds[0];
    var imageHeight = imageBounds[3] - imageBounds[1];

    // 计算缩放比例以使图片的较短边适应路径
    var widthRatio = pathWidth / imageWidth;
    var heightRatio = pathHeight / imageHeight;
    var scaleRatio = Math.max(widthRatio, heightRatio);

    // 使用正确的缩放因子调整放置项目的大小
    placedItem.resize(scaleRatio * 100, scaleRatio * 100, true, true, true, true, scaleRatio * 100, Transformation.DOCUMENTORIGIN);

    // 将图片居中放置在路径内
    var newImageBounds = placedItem.geometricBounds;
    var newImageWidth = newImageBounds[2] - newImageBounds[0];
    var newImageHeight = newImageBounds[3] - newImageBounds[1];

    var deltaX = (pathWidth - newImageWidth) / 2;
    var deltaY = (pathHeight - newImageHeight) / 2;

    placedItem.position = [pathBounds[0] + deltaX, pathBounds[1] + deltaY];

    if (item.typename === "GroupItem" && item.clipped) {
        // 如果是现有的剪切组，替换原来的图片
        var oldImage = item.pageItems[item.pageItems.length - 1];
        if (oldImage.typename === "PlacedItem") {
            oldImage.remove();
        }
        placedItem.move(item, ElementPlacement.PLACEATEND);
    } else {
        // 对于复合路径和普通路径
        placedItem.move(item, ElementPlacement.PLACEBEFORE);
        item.zOrder(ZOrderMethod.BRINGTOFRONT);
        app.selection = null;
        item.selected = true;
        placedItem.selected = true;
        app.executeMenuCommand('makeMask');
    }
}

// 进度条类
function ProgressBar(title, max) {
    this.win = new Window("palette", title, undefined, {closeButton: false});
    this.win.progressBar = this.win.add("progressbar", undefined, 0, max);
    this.win.progressBar.preferredSize.width = 300;
    this.win.show();

    this.update = function(val) {
        this.win.progressBar.value = val;
        this.win.update();
    }

    this.close = function() {
        this.win.close();
    }
}

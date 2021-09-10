---
layout:     post
title:      "SpreadJS结合Springboot实现跨文档数据公式关联（一）"
subtitle:   "———— 原理、需求和思路"
date:       2020-09-29
author:     "Kevin"
header-img: "post-bg.jpg"
tags:
    - SpreadJS
---

> 版主（博主曾任葡萄城技术社区SpreadJS板块版主）在日常沟通、技术社区等与使用SpreadJS的小伙伴们交流的过程中，遇到了很多前后端结合的问题。其实从定位来看，SpreadJS无论从技术架构，还是产品设计层面，都是纯前端的控件。但它毕竟最终还是要结合后端环境来开发部署和应用的，所以我有了写这篇文章的想法。
> 本文利用了一个比较贴近用户的前后端结合的示例工程，来给大家演示一下SpreadJS的一些使用场景和实现技巧，希望能够帮助大家打开思路，能够更好地解决业务问题，提升用户体验。
> 由于涉及的功能点较多，代码量也不小，这个主题我会分为两篇文章来发布，本篇文章是这个系列的第一篇。

**需求背景描述：**
SpreadJS是纯前端控件，受到JS前端语言运行环境的限制，无法原生实现跨Workbook的数据引用。但我们可以通过前后端结合的二次开发，结合业务场景来实现我们的需求。

**需求/功能描述：**
1、 能够远程加载服务器端的ssjson、Excel文档到前端页面；
2、 可以分析表单中跨Workbook的公式引用，手动建立引用目标文档。当多人同时打开相关联的几个文档时，能够远程关联公式的数据（协同的一种）；
3、 能够建立单元格级别的数据关联，可以手动指定关联的目标文档的单元格；
4、 建立好关联单元格后，单元格内右侧出现 ＝、≠的符号，当关联的两个单元格的值相等时显示＝，不相等时显示≠；
5、 同步数据，自动将引用数据回填到引用单元格中。

**实现思路：**
借助Ajax的帮助，上传、下载文件流已经可以完美解决远程文档加载和上传的问题（具体方法，参考前一篇博客《SpreadJS + Vue + Axios 实现服务端加载、上传Excel》）。
静态引用某个后端模板的值、或者是服务器端的数据，其实非常容易理解和实现。原理很简单，示例中采用了命名信息+自定义单元格的方式来实现，一方面实现了需求，另一方面也提供了良好的用户体验。

**运行环境：**
Java SpringBoot+ SpreadJS

**运行方法：**
解压后导入Eclipse（或其它IDE工具），Run Java Application ，访问localhost:8080即可。

**关键代码讲解：**
限于篇幅，本文只讲单元格与后台文档执行关联、删除、同步数据的实现。远程同步、跨Workbook公式的实现，请参考第二篇：SpreadJS结合Springboot实现跨文档数据公式关联（二）。
```js
function CustomValidateCell(hostMargin) {
        this.typeName = "CustomValidateCell";
    this.hostMargin = hostMargin;
        this.margin = 0;
        this.color = 'red';
        this.size = 20;
        this.sideLength = 500;
        this.eq = "lib/eq.png";
        this.neq = "lib/neq.png";
        this.unknown = "lib/unknown.png";
        this.backgroundImage = this.unknown;
}
CustomValidateCell.prototype = new GC.Spread.Sheets.CellTypes.Text();
// 渲染方法，用于实现＝和≠的图标渲染（背景logo）
CustomValidateCell.prototype.paint = function (ctx, value, x, y, w, h, style, options) {
    style.hAlign = GC.Spread.Sheets.HorizontalAlign.left;
    GC.Spread.Sheets.CellTypes.Text.prototype.paint.call(this, ctx, value, x, y, w, h, style, options);
        
    if (!ctx) {
        return;
    }
    var tag = options.sheet.getTag(options.row, options.col, options.sheetArea);
    var startX = x + w - this.size - this.margin;
    var startY = y + (h - this.size) / 2 - this.margin;
    style.backgroundImage = this.backgroundImage;
    GC.Spread.Sheets.CellTypes.Text.prototype.paint.call(this, ctx, "", startX, startY, this.size, this.size, style, options);
};
// 监听鼠标动作和位置
CustomValidateCell.prototype.getHitInfo = function (x, y, cellStyle, cellRect, context) {
    var info = {
        x: x,
        y: y,
        row: context.row,
        col: context.col,
        cellStyle: cellStyle,
        cellRect: cellRect,
        sheetArea: context.sheetArea
    };
    var hitX = x;
    var hitY = y;
    x = cellRect.x;
    y = cellRect.y;
    var w = cellRect.width;
    var h = cellRect.height;

    var startX = x + w - this.size + this.margin;
    var startY = y + (h - this.size) / 2 + this.margin;
    var endX = x + w - this.margin;
    var endY = y + (h + this.size) / 2 - this.margin;
    // 这里判断逻辑参考paint中绘制的逻辑
    if (hitX > startX && hitX < endX && hitY > startY && hitY < endY) {
        info.isReservedLocation = true;
    }
    return info;
};
// 添加鼠标移入事件，弹窗显示引用信息
CustomValidateCell.prototype.processMouseEnter = function (hitInfo) {
    var sheet = hitInfo.sheet;
    if (sheet && hitInfo.isReservedLocation) {
        if (!this._toolTipElement) {
            var div = document.createElement("div");
            $(div).css("position", "absolute")
                .css("border", "1px #C0C0C0 solid")
                .css("box-shadow", "1px 2px 5px rgba(0,0,0,0.4)")
                .css("font", "9pt Arial")
                .css("background", "white")
                .css("padding", 5)
                                .attr("class","toolTipElement");

            this._toolTipElement = div;
        }
        validateCellForCustomerCell(hitInfo.row, hitInfo.col, this);
        $(this._toolTipElement).html($('#validateCellInfo').html())
            .css("top", hitInfo.y + this.hostMargin.y + 15)
            .css("left", hitInfo.x + this.hostMargin.x + 15);
        $(this._toolTipElement).hide();
        document.body.insertBefore(this._toolTipElement, null);
        $(this._toolTipElement).show("fast");
        return true;
    } else {
        if (this._toolTipElement) {
            $(".toolTipElement").remove();
            this._toolTipElement = null;
            return true;
        }
    }
    return false;
};
// 销毁弹窗
CustomValidateCell.prototype.processMouseLeave = function (hitInfo) {
    var sheet = hitInfo.sheet;
        $(".toolTipElement").remove();
    this._toolTipElement = null;
    return false;
};
```

以上代码，采用自定义单元格，实现了关联数据、显示＝或≠图标，弹窗显示引用信息等功能。如图：
![](001.png)
具体实现细节，请结合注释食用。完整Demo参考附件。

**完整Demo参考：**
[SpreadJS_SpringBoot.zip](SpreadJS_SpringBoot.zip)

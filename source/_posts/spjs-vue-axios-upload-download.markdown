---
layout:     post
title:      "SpreadJS + Vue + Axios 实现服务端加载、上传Excel"
subtitle:   "———— 前后端分离开发的非常实用的场景问题"
date:       2020-09-28
author:     "Kevin"
header-img: "post-bg.png"
tags:
    - SpreadJS
---

> Web端实现Excel最为常见的场景之一，就是怎样跟Server端配合实现文件远程加载和上传。
> 实际上这个问题本身跟SpreadJS没有太大关系，因为SpreadJS本质上是个纯前端离线工具包，对文件的导入、导出也都是blob流级的支持，因此，剩下的问题实际上就是怎么解决一个Excel（或其它格式）的上传和下载。

**需求背景描述：**
有经验的同学，看到文章开始的两句其实就够了，自己动手搜一下，网上有很多相关教程。但奈何现在大家都压力大，开发节奏快，为了让大家直接C + V大法拿来就用，我就自己实现了一版。本篇不讲原理（因为网上太多相关介绍），只甩干货，拿来就用（开玩笑，你不测测敢用？）！

**使用方法：**
附件包含两个zip：
1、GCExcelAndSpreadJS_BigWorkbook.zip 是Java服务端示例，解压后导入VSCode或Eclipse后，
运行GcExcelAndSpreadJsApplication的main方法即可启动。（Maven）
[SpringBoot_Server.zip](SpringBoot_Server.zip)
2、Axios实现上传下载Excel.zip 是Vue + Axios实现的前端工程，解压后执行
```shell
npm install
npm start
```
[SpreadJS_Axios_Vue.zip](SpreadJS_Axios_Vue.zip)
即可成功运行。运行后访问 localhost:3000 即可

注意：以上两个前后端工程默认需要在同一个机器上启动，前端用devServer处理的跨域请求。

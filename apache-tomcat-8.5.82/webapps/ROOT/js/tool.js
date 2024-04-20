//初始化工具栏某些操作
function initTool() {
    let colors = [
        ["rgb(0, 0, 0)", "rgb(67, 67, 67)", "rgb(102, 102, 102)", "rgb(204, 204, 204)", "rgb(217, 217, 217)", "rgb(255, 255, 255)"],
        ["rgb(152, 0, 0)", "rgb(255, 0, 0)", "rgb(255, 153, 0)", "rgb(255, 255, 0)", "rgb(0, 255, 0)", "rgb(0, 255, 255)", "rgb(74, 134, 232)", "rgb(0, 0, 255)", "rgb(153, 0, 255)", "rgb(255, 0, 255)"],
        ["rgb(230, 184, 175)", "rgb(244, 204, 204)", "rgb(252, 229, 205)", "rgb(255, 242, 204)", "rgb(217, 234, 211)", "rgb(208, 224, 227)", "rgb(201, 218, 248)", "rgb(207, 226, 243)", "rgb(217, 210, 233)", "rgb(234, 209, 220)", "rgb(221, 126, 107)", "rgb(234, 153, 153)", "rgb(249, 203, 156)", "rgb(255, 229, 153)", "rgb(182, 215, 168)",
            "rgb(162, 196, 201)", "rgb(164, 194, 244)", "rgb(159, 197, 232)", "rgb(180, 167, 214)", "rgb(213, 166, 189)", "rgb(204, 65, 37)", "rgb(224, 102, 102)", "rgb(246, 178, 107)", "rgb(255, 217, 102)", "rgb(147, 196, 125)",
            "rgb(118, 165, 175)", "rgb(109, 158, 235)", "rgb(111, 168, 220)", "rgb(142, 124, 195)", "rgb(194, 123, 160)", "rgb(166, 28, 0)", "rgb(204, 0, 0)", "rgb(230, 145, 56)", "rgb(241, 194, 50)", "rgb(106, 168, 79)",
            "rgb(69, 129, 142)", "rgb(60, 120, 216)", "rgb(61, 133, 198)", "rgb(103, 78, 167)", "rgb(166, 77, 121)", "rgb(91, 15, 0)", "rgb(102, 0, 0)", "rgb(120, 63, 4)", "rgb(127, 96, 0)", "rgb(39, 78, 19)", "rgb(12, 52, 61)", "rgb(28, 69, 135)", "rgb(7, 55, 99)", "rgb(32, 18, 77)", "rgb(76, 17, 48)"]
    ];
    //颜色插件配置
    let colorOpt = {
        allowEmpty: true,
        showInput: true,
        containerClassName: "full-spectrum",
        showInitial: true,
        showPalette: true,
        showSelectionPalette: true,
        showAlpha: true,
        maxPaletteSize: 10,
        clickoutFiresChange: false,
        preferredFormat: "hex8",
        move: function (color) {

        },
        change: function (color) {
            let hexColor = "transparent";
            if (color) {
                hexColor = color.toString();
                let alpha = hexColor.substring(1, 3);
                let trueColor = hexColor.substring(3, 9);
                let col = '#' + trueColor + alpha;
                changeColor($(this), col, true, color);
                $(this).find('#lineColorDiv').css("background-color", '#' + trueColor);
                $(this).find('#lineColorText').text('#' + trueColor);
            }
        },
        beforeShow: function (color) {
            if (color == null) {
                $(this).spectrum('set', '#ffffffff');
            } else {
                let co = '';
                let id = $(this).attr('id');
                if ('chartColor' == id || 'linecolor' == id || 'linecolor1' == id || 'bkcolor' == id || 'bkcolor1' == id || 'bkcolor2' == id || 'color' == id || 'color1' == id) {
                    co = $(this).css('background-color');
                } else {
                    co = $(this).css('border-bottom-color');
                }
                if (co) {
                    $(this).spectrum('set', co);
                }
            }

        },
        hide: function () {
            $("#mask").hide();
        },
        RgbaToHex: function (val) {
            var reg = /^#([0-9a-fA-f]{3}|[0-9a-fA-f]{8})$/;
            if (/^(rgba|RGBA)/.test(val)) {
                var aColor = val.replace(/(?:\(|\)|rgba|RGBA)*/g, "").split(",");
                var strHex = "#";
                for (var i = 0; i < aColor.length; i++) {
                    if (i == 3) {
                        var hex = colorOpt.alphaToHex(Number(aColor[i]));
                        strHex += hex;
                    } else {
                        var hex = Number(aColor[i]).toString(16);
                        if (hex === "0") {
                            hex += hex;
                        }
                        strHex += hex;
                    }

                }
                return strHex;
            } else if (reg.test(val)) {
                var aNum = val.replace(/#/, "").split("");
                if (aNum.length === 6) {
                    return val;
                } else if (aNum.length === 3) {
                    var numHex = "#";
                    for (var i = 0; i < aNum.length; i += 1) {
                        numHex += (aNum[i] + aNum[i]);
                    }
                    return numHex;
                }
            } else {
                return val;
            }
        },
        alphaToHex: function (val) {
            let num = val * 255;
            if (num == 0) {
                return '00';
            } else {
                let value = (Math.round(num)).toString(16);
                if (value.length == 1) {
                    value = '0' + value;
                }
                return value;
            }
        },
        palette: colors
    };

    $("img[name='bgcolor']").spectrum(colorOpt);
    $("img[name='ftcolor']").spectrum(colorOpt);
    $("#linecolor").spectrum(colorOpt);

    layui.use('upload', function () {
        let upload = layui.upload;
        upload.render({
            elem: '#importTempl'
            , url: base + '/designSys/getJsonData'
            , accept: 'file'
            , exts: 'xlsx|cel'
            , done: function (res) {
            }
            , before: function (obj) {
                let files = obj.pushFile();
                obj.preview(function (index, file, result) { //得到文件base64编码，比如图片
                    let fileName = file.name;
                    if(fileName.endsWith('.cel')){
                        let ts = result.indexOf('base64,') + 7;
                        let tt = result.substr(ts, result.length);
                        let str = new Base64().decode(tt);
                        let name = file.name.substr(0, file.name.indexOf('.cel'));
                        spread.restoreFromJsonStream(str);
                    }else{
                        let ts = result.indexOf('base64,') + 7;
                        let t = result.substr(ts, result.length);
                        let name = file.name.substr(0, file.name.lastIndexOf('.'));
                        spread.importXlsxStream(t);
                    }

                });
            }
        });
    });

    layui.use('upload', function () {//插入图片
        var upload = layui.upload;
        upload.render({ //允许上传的文件后缀
            elem: '#addImg'
            , url: base + '/designSys/getJsonData'
            , accept: 'images'
            , done: function (res) {
            }
            , before: function (obj) {
                obj.preview(function (index, file, result) {
                    var t = result.substr(result.indexOf('base64,') + 7, result.length);
                    spread.setSelCellBkPic(t);
                    let a = spread.isSelCellAdaptImageSize();
                    var b = spread.isSelCellRawImageScale();
                    $("#imgsf").prop("checked", a == 1 ? true : false);
                    $("#imgkeep").prop("checked", b == 1 ? true : false);
                    if (a) {
                        $("#imgkeep").removeAttr('disabled');
                    }
                });
            }
        });
    });

    //初始化字号
    for (let i = 1; i < 100; i++) {
        let opt;
        if (i == 12) {
            opt = '<option value="' + i + '" selected>' + i + '</option>';
        } else {
            opt = '<option value="' + i + '">' + i + '</option>';
        }
        $("select[name='fontsize']").append(opt);
        $("select[name='fontsizes']").append(opt);
    }
    //默认字号大小为18
    //$("select[name='fontsize']").val(18).attr("selected", true);
    //初始化边框操作
    $('#border').bind('click', function (event) {
        event.stopPropagation(); //阻止冒泡，不响应body的点击事件
        let list = $('#borderDiv');
        let X = $(this).offset().top + 20;
        let Y = $(this).offset().left;
        list.css({"top": X, "left": Y, "z-index": 100}).show();
    })
    //borderDiv中边框按钮操作
    $('img[name="borders"]').bind('click', function () {
        //点击值
        let val = $(this).attr("attr");
        let color = $("#lineColorText").text(); //当前选中的颜色
        let penStyle = $("#ckline").children(0).children(0).attr('val'); //
        $(this).siblings().removeClass('choose');
        $(this).addClass('choose');
        if (val == '00') { //取消所有样式
            spread.cancelSelCellLineStyle();
        } else if (val == '7') {//加粗边框
            //spread.setSelCellLineStyle(5, 1, 1, '#000000');
            spread.setSelCellLineStyle(5, penStyle, 1, color);
        } else {
            spread.setSelCellLineStyle(val, penStyle, 0, color);
        }
        //修改边框样式图案
        $('#border').attr('src', $(this).attr('src'));
        $('#borderDiv').hide();

    })

    //单元格位置
    $('#cellPos').unbind().bind('keypress', function (event) {
        if (event.keyCode == '13') { //回车键
            let val = $(this).val();
            let value = spread.moveSelFrameToCell(val); //跳转到指定的位置
            let cellVal = spread.getSelCellText();
            $('#editArea').val(cellVal);
        }
    });
    //单元格内容
    $('#editArea').bind('keyup', function () { //单元格文本修改时，设计器上方的编辑框对应修改
        let val = $(this).val();
        spread.setSelCellText(val);
    });

    // 图片水平对齐
    $("img[name='imghAlign']").click(function () {
        $("img[name='imghAlign']").each(function () {
            $(this).removeClass('choose');
        });
        $(this).addClass('choose');
        spread.setSelCellImageAlignX($(this).attr('val'));
    });

    // 图片垂直对齐
    $("img[name='imgvAlign']").click(function () {
        $("img[name='imgvAlign']").each(function () {
            $(this).removeClass('choose');
        });
        $(this).addClass('choose');
        spread.setSelCellImageAlignY($(this).attr('val'));
    });

    //图片缩放
    $("#imgsf").change(function () {
        let val = $(this).prop('checked');
        spread.setSelCellAdaptImageSize(val);
        if (val) {
            //保持原比例
            $("#imgkeep").removeAttr('disabled');
            $('#tz2_tb').find('img').attr('disabled', true);
        } else {
            $('#tz2_tb').find('img').removeAttr('disabled');
            $("#imgkeep").attr('disabled', true);
        }
    });
    //保持原比例
    $("#imgkeep").change(function () {
        spread.setSelCellRawImageScale($(this).prop('checked'));
    });

    $('.cellType').find('div').bind('click', function () {
        $(this).addClass('choose1');
        $(this).siblings().removeClass('choose1');
        let cla = $(this).attr('id');
        if (cla == 'normal-t') {
            $('.formatType').children('div').hide();
            $('.formatType').find('.normal').show();
        } else if (cla == 'number-t') {
            $('.formatType').children('div').hide();
            $('.formatType').find('.number').show();
            $('.number').find('.number-data').removeClass('choose1');
            $('.number').find('.number-data').eq(0).addClass('choose1');
        } else if (cla == 'date-t') {
            $('.formatType').children('div').hide();
            $('.formatType').find('.date').show();
            $('.date').find('.number-data').removeClass('choose1');
            $('.date').find('.number-data').eq(0).addClass('choose1');
        } else if (cla == 'time-t') {
            $('.formatType').children('div').hide();
            $('.formatType').find('.time').show();
            $('.time').find('.number-data').removeClass('choose1');
            $('.time').find('.number-data').eq(0).addClass('choose1');
        } else if (cla == 'money-t') {
            $('.formatType').children('div').hide();
            $('.formatType').find('.money').show();
            $('.money').find('.number-data').removeClass('choose1');
            $('.money').find('.number-data').eq(0).addClass('choose1');
        } else if (cla == 'percent-t') {
            $('.formatType').children('div').hide();
            $('.formatType').find('.percent').show();
            $('.percent').find('.number-data').removeClass('choose1');
            $('.percent').find('.number-data').eq(0).addClass('choose1');
        } else {
            $('.formatType').children('div').hide();
            $('.formatType').find('.code').show();
            $('.code').find('.number-data').removeClass('choose1');
            $('.code').find('.number-data').eq(0).addClass('choose1');
        }
    })

    $('.number-data').bind('click', function () {
        $(this).addClass('choose1');
        $(this).siblings().removeClass('choose1');
    })

    $('#sheetMenu a').click(function (event) {
        var attr = $(this).children('img').attr('attr');
        if ('delSh' == attr) {
            spread.removeCurrentSheet();
            $('#sheetMenu').hide();
        } else if ('setNameSh' == attr) { //修改sheet名称
            $('#sheetMenu').hide();
            var index = layer.open({
                type: 2,
                area: ['350px', '180px'],
                closeBtn: 0,
                resize: false,
                title: ['sheet名称', 'height:30px;line-height:30px'],
                content: ['pages/sheetName.html', 'no'],
                btn: ['确定', '关闭'],
                btnAlign: 'c',
                end: function () {
                    $('#mask').trigger("click");
                },
                yes: function (index, layero) {
                    var iframeWin = window[layero.find('iframe')[0]['name']];
                    var res = iframeWin.getPage();
                    if (res.newVal != '') {
                        if (res.oldVal != res.newVal) {
                            var t = spread.setCurrentSheetName(res.newVal);
                            if (t == -2) {
                                layer.alert('名称已经存在');
                            } else {
                                layer.close(index);
                            }
                        } else {
                            layer.close(index);
                        }
                    } else {
                        layer.alert('请输入sheet名称');
                    }
                },
                success: function (layero, index) {
                    var iframeWin = window[layero.find('iframe')[0]['name']];
                    iframeWin.setVal(spread.getCurrentSheetName() , index);
                }
            });
        } else if ('addSh' == attr) {
            $('#mask').trigger("click");
            spread.appendSheet();
            $('#sheetMenu').hide();
        } else if ('copySheet' == attr) {
            $('#sheetMenu').hide();
            var index = layer.open({
                type: 2,
                area: ['350px', '180px'],
                closeBtn: 0,
                resize: false,
                title: ['sheet名称', 'height:30px;line-height:30px'],
                content: ['pages/sheetName.html', 'no'],
                btn: ['确定', '关闭'],
                btnAlign: 'c',
                end: function () {
                    $('#mask').trigger("click");
                },
                yes: function (index, layero) {
                    var iframeWin = window[layero.find('iframe')[0]['name']];
                    var name = iframeWin.getPage();
                    if (name.newVal != '') {
                        //拷贝当前sheet
                        var t = spread.cloneSheetByName(name.newVal);
                        if (t == -2) {
                            layer.alert('拷贝失败,sheet名已存在');
                        }
                        layer.close(index);
                    } else {
                        layer.alert('请输入sheet名称');
                    }
                },
                success: function (layero, index) {
                    var iframeWin = window[layero.find('iframe')[0]['name']];
                    iframeWin.setVal(spread.getCurrentSheetName());
                }
            });

        }
    });

    //边框样式
    $("#ckline").click(function () {
        let X = $(this).children(0).offset().top + 25;
        let Y = $(this).children(0).offset().left;
        $("#lines").css({"top": X, "left": Y ,"z-index":10001}).show();
        //$("#lines").focus();
        $("#lines").blur(function () {
            $(this).hide();
        });
    });

    //选择样式事件
    $("div[name='lineval']").click(function () {
        let val = $(this).attr('val'),
            url = 'image/line/l' + val + '.png';
        let img = $("#ckline").children(0).children(0);
        img.attr('src', url);
        img.attr('val', val);
        let img1 = $("#ckline1").children(0).children(0);
        img1.attr('src', url);
        img1.attr('val', val);
        $("#lines").hide();
    });

}


function changeColor(obj, hexColor, ct, color) {
    var url;
    if (obj.attr("name") == "bgcolor") { //背景颜色
        url = "images/design/42_1.png";
        spread.setSelCellBKColor(hexColor);
        obj.css("border-bottom", "4px solid " + hexColor + "");
    } else if (obj.attr("name") == "ftcolor") { //文本颜色
        url = "images/design/43_1.png";
        spread.setSelCellFontColor(hexColor);
        //obj.attr("src", url);
        obj.css("border-bottom", "4px solid " + hexColor + "");
    }
}

function importCel() {

}

function saveC() {
    let content = spread.getContent();
    let blob = new Blob([content], {type: "application/json"});
    saveAs(blob, "template.cel");
}

function copyS() {
    spread.copy();
}

function pasteS() {
    setTimeout(function () {
        let content = $("#ef_copy_area").val();
        let str = spread.encode(content);
        spread.copyToSheet(str);
        spread.paste();
    }, 100)
}

function cutS() {
    spread.cut();
}

function undoC() {
    spread.undo();
}

function comebackC() {
    spread.comeback();
}

function formatC() {
    spread.formatBrush();
}

function mergeCell() {
    spread.mergeCell();
}

function splitCell() {
    spread.splitCell();
}

function appendRow() {
    spread.appendRow();
}

function insertRow() {
    spread.insertRow();
}

function deleteRow() {
    spread.deleteRow();
}

function appendColumn() {
    spread.appendColumn();
}

function insertColumn() {
    spread.insertColumn();
}

function deleteColumn() {
    spread.deleteColumn();
}

//左对齐
function alignLeft() {
    $('img[name="left"]').addClass('choose');
    $('img[name="center"]').removeClass('choose');
    $('img[name="right"]').removeClass('choose');
    spread.setSelCellAlignH(1);
}

//居中
function alignCenter() {
    $('img[name="center"]').addClass('choose');
    $('img[name="left"]').removeClass('choose');
    $('img[name="right"]').removeClass('choose');
    spread.setSelCellAlignH(4);
}

//右对齐
function alignRight() {
    $('img[name="right"]').addClass('choose');
    $('img[name="center"]').removeClass('choose');
    $('img[name="left"]').removeClass('choose');
    spread.setSelCellAlignH(2);
}

//上对齐
function alignTop() {
    $('img[name="top"]').addClass('choose');
    $('img[name="middle"]').removeClass('choose');
    $('img[name="bottom"]').removeClass('choose');
    spread.setSelCellAlignV(16);
}

//垂直对齐
function alignMiddle() {
    $('img[name="middle"]').addClass('choose');
    $('img[name="top"]').removeClass('choose');
    $('img[name="bottom"]').removeClass('choose');
    spread.setSelCellAlignV(64);
}

//下对齐
function alignBottom() {
    $('img[name="bottom"]').addClass('choose');
    $('img[name="middle"]').removeClass('choose');
    $('img[name="top"]').removeClass('choose');
    spread.setSelCellAlignV(32);
}

function boldC() {
    let font = spread.getSelCellFont(); //获取字体属性
    let fontJson = JSON.parse(font);
    let bold = fontJson.bold; //是否加粗
    if (bold) {
        spread.setSelCellFontBold(false);
        $('img[name="bold"]').removeClass('choose');
    } else {
        spread.setSelCellFontBold(true);
        spread.setSelCellFontWeight(700);
        $('img[name="bold"]').addClass('choose');
    }

}

//斜体
function italicC() {
    let font = spread.getSelCellFont(); //获取字体属性
    let fontJson = JSON.parse(font);
    let italic = fontJson.italic;
    if (!italic) { //斜体
        $('img[name="italic"]').addClass('choose');
    } else {
        $('img[name="italic"]').removeClass('choose');
    }
    spread.setSelCellFontItalic(!italic);

}

//下划线
function underLineC() {
    let font = spread.getSelCellFont(); //获取字体属性
    let fontJson = JSON.parse(font);
    let underline = fontJson.underline; //是否加粗
    if (underline) {
        spread.setSelCellFontUnderline(false);
        $('img[name="underline"]').removeClass('choose');
    } else {
        spread.setSelCellFontUnderline(true);
        $('img[name="underline"]').addClass('choose');
    }
}

function changeFontSize(obj) {
    spread.setSelCellFontPointSize($(obj).val());
}

function adaptLineHeightC() {
    let isAdapt = spread.isSelCellAdaptTextHeight(); //是否自适应行高
    if (isAdapt) { //自适应
        $('img[name="adapt"]').removeClass('choose');
        spread.setSelCellAdaptTextHeight(false);
    } else {
        $('img[name="adapt"]').addClass('choose');
        $('img[name="adaptColumn"]').removeClass('choose');
        spread.setSelCellAdaptTextHeight(true);
    }
}

function adaptC() {
    let isAdapt = spread.isSelCellAdaptTextHeight(); //是否自适应行高
    if (isAdapt) { //自适应
        $('img[name="adapt"]').removeClass('choose');
        spread.setSelCellAdaptTextHeight(false);
    } else {
        $('img[name="adapt"]').addClass('choose');
        $('img[name="adaptColumn"]').removeClass('choose');
        spread.setSelCellAdaptTextHeight(true);
    }
}

function adaptColumnC() {
    let isAdapt = spread.isSelCellAdaptTextWidth(); //是否自适应列宽
    if (isAdapt) {
        $('img[name="adaptColumn"]').removeClass('choose');
        spread.setSelCellAdaptTextWidth(!isAdapt);
        //canvasEvent.Cell.setSelCellAdaptTextHeight(isAdapt);
    } else {
        $('img[name="adaptColumn"]').addClass('choose');
        $("img[name='adapt']").removeClass('choose');
        spread.setSelCellAdaptTextWidth(!isAdapt);
        spread.setSelCellAdaptTextHeight(isAdapt);
    }
}

function changeFontFamily(obj) {
    let font = $(obj).val();
    spread.setSelCellFontFamily(font);
}

function initCellProp() {

    let text = spread.getSelCellText(); //获取当前单元格文本
    $('#editArea').val(text); //设置编辑框值

    let cellPos = spread.getCellPos();
    $('#cellPos').val(cellPos);

    /**
     * 单元格数据开始
     * **/

    //边框图片恢复原始状态
    $('#border').attr("src", "image/line.png");
    $('img[name="borders"]').removeClass('choose');

    let font = spread.getSelCellFont(); //获取字体属性
    let fontJson = JSON.parse(font);
    if (fontJson.bold) {
        $('img[name="bold"]').addClass('choose');
    } else {
        $('img[name="bold"]').removeClass('choose');
    }
    if (fontJson.italic) {
        $('img[name="italic').addClass('choose');
    } else {
        $('img[name="italic"]').removeClass('choose');
    }
    if (fontJson.underline) {
        $('img[name="underline').addClass('choose');
    } else {
        $('img[name="underline"]').removeClass('choose');
    }
    $('select[name="fontfamily"]').val(fontJson.family).prop("selected", true);//字体
    $('select[name="fontsize"]').val(fontJson.pointSize).prop("selected", true);//字体大小

    let color = spread.getSelCellFontColor(); //文本颜色
    //$("img[name='ftcolor']").attr("src", "images/design/43_1.png");
    $("img[name='ftcolor']").css("border-bottom", "4px solid " + color);

    let bkColor = spread.getSelCellBKColor(); //背景颜色
    $("img[name='bgcolor']").css("border-bottom", "4px solid " + bkColor);

    let alignH = spread.getSelCellAlignH(); //水平方向
    if (alignH == 1) {//left
        $('img[name="left"]').addClass('choose');
        $('img[name="center"]').removeClass('choose');
        $('img[name="right"]').removeClass('choose');
    } else if (alignH == 4) {//center
        $('img[name="center"]').addClass('choose');
        $('img[name="left"]').removeClass('choose');
        $('img[name="right"]').removeClass('choose');
    } else if (alignH == 2) {//right
        $('img[name="right"]').addClass('choose');
        $('img[name="left"]').removeClass('choose');
        $('img[name="center"]').removeClass('choose');
    } else {
        $('img[name="left"]').removeClass('choose');
        $('img[name="center"]').removeClass('choose');
        $('img[name="right"]').removeClass('choose');
    }

    let alignV = spread.getSelCellAlignV(); //水平方向
    if (alignV == 16) {//top
        $('img[name="top"]').addClass('choose');
        $('img[name="middle"]').removeClass('choose');
        $('img[name="bottom"]').removeClass('choose');
    } else if (alignV == 64) {//center
        $('img[name="middle"]').addClass('choose');
        $('img[name="top"]').removeClass('choose');
        $('img[name="bottom"]').removeClass('choose');
    } else if (alignV == 32) {//right
        $('img[name="bottom"]').addClass('choose');
        $('img[name="top"]').removeClass('choose');
        $('img[name="middle"]').removeClass('choose');
    } else {
        $('img[name="top"]').removeClass('choose');
        $('img[name="middle"]').removeClass('choose');
        $('img[name="bottom"]').removeClass('choose');
    }

    let adapt = spread.isSelCellAdaptTextHeight();
    if (adapt) {
        $('img[name="adapt"]').addClass('choose');
    } else {
        $('img[name="adapt"]').removeClass('choose');
    }

    let dataType = spread.getFieldDataType(); //获取数据类型
    $('#dstype').val(dataType).prop("selected", true);

}

function showType() {
    let width = $('#export').width();
    let height = $('#export').height();
    $('#typeDiv').css({
        top:height + 'px',
        left: width/2 + 'px'
    }).show();
}

function exportFile(type) {
    $('#typeDiv').hide();
    let sval = spread.getContent();//临时保存
    let temp = 'file';
    let blob = new Blob([sval], {type: 'application/json'});
    let formdata = new FormData();
    formdata.append('file', blob);
    formdata.append('fileName', temp);
    $.ajax({
        //生成臨時文件
        url: base + '/thirdSys/generateFile',
        type: 'post',
        processData: false,
        contentType: false,
        dataType: "json",
        data: formdata,
        success: function (data) {
            if (data.state == 'success') {
                let name = data.name;
                if (type == 1) {
                    window.open(base + '/thirdSys/exportDesignPdf?name=' + name);
                } else if (type == 2) {
                    window.open(base + '/thirdSys/exportDesignExcel?name=' + name);
                } else {
                    window.open(base + '/thirdSys/exportDesignWord?name=' + name);
                }
            } else {

            }
        },
        error: function (info) {
            layui.use('layer', function () {
                var layer = layui.layer;
            });
            layer.msg('此功能需要搭建盈帆报表后台服务方可正常使用');
        }
    });

}

function Base64() {

    // private property
    var _keyStr = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/=";

    // public method for encoding
    this.encode = function (input) {
        var output = "";
        var chr1, chr2, chr3, enc1, enc2, enc3, enc4;
        var i = 0;
        input = _utf8_encode(input);
        while (i < input.length) {
            chr1 = input.charCodeAt(i++);
            chr2 = input.charCodeAt(i++);
            chr3 = input.charCodeAt(i++);
            enc1 = chr1 >> 2;
            enc2 = ((chr1 & 3) << 4) | (chr2 >> 4);
            enc3 = ((chr2 & 15) << 2) | (chr3 >> 6);
            enc4 = chr3 & 63;
            if (isNaN(chr2)) {
                enc3 = enc4 = 64;
            } else if (isNaN(chr3)) {
                enc4 = 64;
            }
            output = output +
                _keyStr.charAt(enc1) + _keyStr.charAt(enc2) +
                _keyStr.charAt(enc3) + _keyStr.charAt(enc4);
        }
        return output;
    };

    // public method for decoding
    this.decode = function (input) {
        var output = "";
        var chr1, chr2, chr3;
        var enc1, enc2, enc3, enc4;
        var i = 0;
        input = input.replace(/[^A-Za-z0-9\+\/\=]/g, "");
        while (i < input.length) {
            enc1 = _keyStr.indexOf(input.charAt(i++));
            enc2 = _keyStr.indexOf(input.charAt(i++));
            enc3 = _keyStr.indexOf(input.charAt(i++));
            enc4 = _keyStr.indexOf(input.charAt(i++));
            chr1 = (enc1 << 2) | (enc2 >> 4);
            chr2 = ((enc2 & 15) << 4) | (enc3 >> 2);
            chr3 = ((enc3 & 3) << 6) | enc4;
            output = output + String.fromCharCode(chr1);
            if (enc3 != 64) {
                output = output + String.fromCharCode(chr2);
            }
            if (enc4 != 64) {
                output = output + String.fromCharCode(chr3);
            }
        }
        output = _utf8_decode(output);
        return output;
    };

    // private method for UTF-8 encoding
    _utf8_encode = function (string) {
        string = string.replace(/\r\n/g, "\n");
        var utftext = "";
        for (var n = 0; n < string.length; n++) {
            var c = string.charCodeAt(n);
            if (c < 128) {
                utftext += String.fromCharCode(c);
            } else if ((c > 127) && (c < 2048)) {
                utftext += String.fromCharCode((c >> 6) | 192);
                utftext += String.fromCharCode((c & 63) | 128);
            } else {
                utftext += String.fromCharCode((c >> 12) | 224);
                utftext += String.fromCharCode(((c >> 6) & 63) | 128);
                utftext += String.fromCharCode((c & 63) | 128);
            }

        }
        return utftext;
    };

    // private method for UTF-8 decoding
    _utf8_decode = function (utftext) {
        var string = "";
        var i = 0;
        var c = c1 = c2 = 0;
        while (i < utftext.length) {
            c = utftext.charCodeAt(i);
            if (c < 128) {
                string += String.fromCharCode(c);
                i++;
            } else if ((c > 191) && (c < 224)) {
                c2 = utftext.charCodeAt(i + 1);
                string += String.fromCharCode(((c & 31) << 6) | (c2 & 63));
                i += 2;
            } else {
                c2 = utftext.charCodeAt(i + 1);
                c3 = utftext.charCodeAt(i + 2);
                string += String.fromCharCode(((c & 15) << 12) | ((c2 & 63) << 6) | (c3 & 63));
                i += 3;
            }
        }
        return string;
    }
}


function editComment() {
    let index = layer.open({
        type: 2,
        area: ['400px', '300px'],
        closeBtn: 0,
        resize: false,
        title: ['单元格注释', 'height:30px;line-height:30px'],
        content: ['pages/comment.html', 'no'],
        btn: ['确定', '关闭'],
        btnAlign: 'c',
        end: function () {

        },
        yes: function (index, layero) {
            let iframeWin = window[layero.find('iframe')[0]['name']];
            let expr = iframeWin.getExpr();
            spread.setSelCellComment(expr) //设置单元格注释;
            layer.close(index);
        },
        success: function (layero, index) {
            let iframeWin = window[layero.find('iframe')[0]['name']];
            let val = spread.getSelCellComment(); //获取单元格注释
            iframeWin.setExpr(val);
        }
    });
}

function delComment() {
    spread.setSelCellComment('');
}

function autoHidden() {
    let flag = spread.getSelCellCommentHide();
    spread.setSelCellCommentHide(!flag)
}

function addOblC() {
    spread.setSelCellsType(9);
}

function fixCell() {
    let index = layer.open({
        type: 2,
        area: ['370px', '250px'],
        closeBtn: 0,
        resize: false,
        title: ['固定行列', 'height:30px;line-height:30px'],
        content: 'pages/fixcell.html',
        btn: ['确定', '关闭'],
        btnAlign: 'c',
        end: function () {

        },
        yes: function (index, layero) {
            let iframeWin = window[layero.find('iframe')[0]['name']];
            let page = iframeWin.getPage();
            let t = spread.setFixedColRow(page.fixColumnCount, page.fixRowCount, page.fixFootRowCount, page.fixFootBeginRow);
            layer.close(index);
        },
        success: function (layero, index) {
            let iframeWin = window[layero.find('iframe')[0]['name']];
            let json = spread.getFixedColRow();
            iframeWin.init(json);
        }
    });
}

function insertFormula() {
    var index = layer.open({
        type: 2,
        area: ['700px', '650px'],
        closeBtn: 0,
        maxmin: true,
        resize: false,
        title: ['公式编辑', 'height:30px;line-height:30px'],
        content: ['pages/expr.html', 'no'],
        btn: ['确定', '关闭'],
        btnAlign: 'c',
        end: function () {

        },
        yes: function (index, layero) {
            let iframeWin = window[layero.find('iframe')[0]['name']];
            let expr = iframeWin.getExpr();
            if (expr.indexOf('=') != 0) {
                expr = "=" + expr;
            }
            spread.setSelCellText(expr);
            layer.close(index);
        },
        success: function (layero, index) {
            let iframeWin = window[layero.find('iframe')[0]['name']];
            let val = spread.getSelCellText();
            let valArr = val.split('');
            if (valArr[0] == '=') {
                val = val.replace('=', '');
            }
            iframeWin.setExpr(val);
        }
    });
}

function setTableRegion() {
    spread.setSelDesignTableRegion(); //设置表格区域
    spread.setSelTableRegion();
}

function removeTableRegion() {
    spread.removeSelDesignTableRegion(); //取消表格区域
    spread.removeSelTableRegion();
}

function filterByText(json) {
    let info = JSON.stringify(json);
    spread.filterTableRegionText(info);
}

function sort(json, order) {
    let info = JSON.stringify(json);
    spread.orderTableRegion(info, order);
    layer.closeAll();
}

function openFilter(title) {
    layer.open({
        type: 2,
        area: ['500px', '300px'],
        closeBtn: 0,
        maxmin: false,
        title: ['自定义筛选方式', 'height:30px;line-height:30px'],
        content: ['pages/numFilter.html', 'no'],
        btn: ['确定', '关闭'],
        resize: false,
        btnAlign: 'c',
        end: function () {

        },
        success: function (layero, i) {
            let iframeWin = window[layero.find('iframe')[0]['name']];
            iframeWin.init(title);
        },
        yes: function (index, layero) {
            let iframeWin = window[layero.find('iframe')[0]['name']];
            let res = iframeWin.getRules();
            let info = JSON.stringify(res);
            spread.filterTableRegionTextByExpr(info);
            layer.closeAll();
        }
    });
}

function openStrFilter(title) {
    layer.open({
        type: 2,
        area: ['500px', '300px'],
        closeBtn: 0,
        maxmin: false,
        title: ['自定义筛选方式', 'height:30px;line-height:30px'],
        content: ['pages/strFilter.html', 'no'],
        btn: ['确定', '关闭'],
        resize: false,
        btnAlign: 'c',
        end: function () {

        },
        success: function (layero, i) {
            let iframeWin = window[layero.find('iframe')[0]['name']];
            iframeWin.init(title);
        },
        yes: function (index, layero) {
            let iframeWin = window[layero.find('iframe')[0]['name']];
            let res = iframeWin.getRules();
            let info = JSON.stringify(res);
            spread.filterTableRegionTextByExpr(info);
            layer.closeAll();
        }
    });
}

function restore(info) {
    spread.restoreShowAllRowsInTableRegion(JSON.stringify(info));
    layer.closeAll();
}

function switchC() {
    let show = spread.isShowExprValue();
    spread.setShowExprValue(!show);
}

function switchR() {
    let show = spread.isShowColRowShadow();
    spread.setShowColRowShadow(!show);
}

function aboutAs(){
    let width = $('body').width();
    let height = $('body').height();
    let eleWidth = $('.popup').width();
    let eleHeight = $('.popup').height();
    let left = (width - eleWidth)/2;
    let top = (height - eleHeight)/2;
    $('.popup').css({
        left:left,
        top:top
    }).show();
}

function closeAbout(){
    $('.popup').hide();
}

function imageSettingC() {
    $('#imageSetting').css({
        'right': '200px',
        'top': '200px',
        'z-index': 10000
    }).show();

    let imgalignH = spread.getSelCellImageAlignX()
    $("img[name='imghAlign']").each(function () {
        if ($(this).attr('val') == imgalignH) {
            $(this).addClass('choose');
        } else {
            $(this).removeClass('choose');
        }
    });
    //图片垂直
    let imgalignV = spread.getSelCellImageAlignY()
    $("img[name='imgvAlign']").each(function () {
        if ($(this).attr('val') == imgalignV) {
            $(this).addClass('choose');
        } else {
            $(this).removeClass('choose');
        }
    });
    //图片缩放
    let isSelCellAdaptImageSize = spread.isSelCellAdaptImageSize();
    if (isSelCellAdaptImageSize) {
        $("#imgkeep").removeAttr('disabled');
    } else {
        $("#imgkeep").attr('disabled', true);
    }
    $("#imgsf").prop("checked", isSelCellAdaptImageSize);
    //保持原比例
    let isSelCellRawImageScale = spread.isSelCellRawImageScale();
    $("#imgkeep").prop("checked", isSelCellRawImageScale);
}

function confirm() {
    let cellType = $('.cellType').find('.choose1').attr('id');
    let valueType;
    if (cellType == 'normal-t') {
        spread.setSelCellsType(0);
    } else if (cellType == 'number-t') {
        spread.setSelCellsType(2);
        let type = $('.number').find('.choose1').attr('datatype');
        valueType = parseInt(type);
    } else if (cellType == 'date-t') {
        spread.setSelCellsType(4);
        let type = $('.date').find('.choose1').attr('datatype');
        valueType = parseInt(type);
    } else if (cellType == 'time-t') {
        spread.setSelCellsType(5)
        let type = $('.time').find('.choose1').attr('datatype');
        valueType = parseInt(type);
    } else if (cellType == 'money-t') {
        spread.setSelCellsType(6);
        let type = $('.money').find('.choose1').attr('datatype');
        valueType = parseInt(type);
    } else if (cellType == 'percent-t') {
        spread.setSelCellsType(8);
        let type = $('.percent').find('.choose1').attr('datatype');
        valueType = parseInt(type);
    } else {
        spread.setSelCellsType(10);
        let type = $('.code').find('.choose1').attr('datatype');
        valueType = parseInt(type);
    }
    //设置数据格式
    if (cellType == 'code-t') {
        spread.setSelCellsBarCodeType(valueType);
    } else {
        spread.setSelCellsStyle(valueType);
    }
    $('.format-container').hide();
}

function cancel() {
    $('.format-container').hide();
}

function dataFormat() {


    $('.format-container').css({
        'right': '200px',
        'top': '200px',
        'z-index': 10000
    }).show();

    let cellType = spread.getSelCellsType(); //单元格类型
    let cellStyle = spread.getSelCellsStyle(); //单元格样式
    if(cellType != 0 && cellType != 1){
        if(cellType == 10){
            $('.formatType').children('div').hide();
            $('.formatType').children('div').eq(6).show();
            $('.cellType').children('div').removeClass('choose1');
            $('#code-t').addClass('choose1');
            let barCode = spread.getSelCellsBarCodeType();
            let codeTypeDiv = $('.code').children('div');
            $.each(codeTypeDiv , function(i,e){
                let dataType = $(e).attr('datatype');
                if(dataType == barCode){
                    $(this).addClass('choose1');
                }else{
                    $(this).removeClass('choose1');
                }
            })

        }else if(cellType == 2){
            $('.formatType').children('div').hide();
            $('.formatType').children('div').eq(1).show();
            $('.cellType').children('div').removeClass('choose1');
            $('#number-t').addClass('choose1');
            let typeDiv = $('.number').children('div');
            $.each(typeDiv , function(i,e){
                let dataType = $(e).attr('datatype');
                if(dataType == cellStyle){
                    $(this).addClass('choose1');
                }else{
                    $(this).removeClass('choose1');
                }
            })
        }else if(cellType == 4){
            $('.formatType').children('div').hide();
            $('.formatType').children('div').eq(2).show();
            $('.cellType').children('div').removeClass('choose1');
            $('#date-t').addClass('choose1');
            let typeDiv = $('.date').children('div');
            $.each(typeDiv , function(i,e){
                let dataType = $(e).attr('datatype');
                if(dataType == cellStyle){
                    $(this).addClass('choose1');
                }else{
                    $(this).removeClass('choose1');
                }
            })
        }else if(cellType == 5){
            $('.formatType').children('div').hide();
            $('.formatType').children('div').eq(3).show();
            let typeDiv = $('.time').children('div');
            $('.cellType').children('div').removeClass('choose1');
            $('#time-t').addClass('choose1');
            $.each(typeDiv , function(i,e){
                let dataType = $(e).attr('datatype');
                if(dataType == cellStyle){
                    $(this).addClass('choose1');
                }else{
                    $(this).removeClass('choose1');
                }
            })
        }else if(cellType == 6){
            $('.formatType').children('div').hide();
            $('.formatType').children('div').eq(4).show();
            let typeDiv = $('.money').children('div');
            $('.cellType').children('div').removeClass('choose1');
            $('#money-t').addClass('choose1');
            $.each(typeDiv , function(i,e){
                let dataType = $(e).attr('datatype');
                if(dataType == cellStyle){
                    $(this).addClass('choose1');
                }else{
                    $(this).removeClass('choose1');
                }
            })
        }else if(cellType == 8){
            $('.formatType').children('div').hide();
            $('.formatType').children('div').eq(5).show();
            let typeDiv = $('.percent').children('div');
            $('.cellType').children('div').removeClass('choose1');
            $('#percent-t').addClass('choose1');
            $.each(typeDiv , function(i,e){
                let dataType = $(e).attr('datatype');
                if(dataType == cellStyle){
                    $(this).addClass('choose1');
                }else{
                    $(this).removeClass('choose1');
                }
            })
        }
    }else{
        $('.cellType').children('div').removeClass('choose1');
        $('#normal-t').addClass('choose1');
        $('.formatType').children('div').hide();
        $('.formatType').children('div').eq(0).show();
    }

}

function changeDsType(obj){
    let val = $(obj).val();
    spread.setFieldDataType(parseInt(val));
}

function websiteC(){
    window.open('http://www.efreport.com');
}




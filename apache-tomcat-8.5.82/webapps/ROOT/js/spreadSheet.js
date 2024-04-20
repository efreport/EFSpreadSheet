/*
 * spreadSheet core
 * v7.5
 *
 *
 * Date: 2024-01-15
 */

(function ($) {
    var settings = {},
        _variables = {
            fontIndex: 0,
            fontLength: 0,
            jsIndex: 0,
            jsLength: 0,
            fontFile: [],
            jsContentMap: {},
            jsDoneFlag: 0,
            timer: ''
        },
        _setting = {
            columnNum: 8, //列数
            rowNum: 10, //行数
            rowHeight: 27, //行高
            columnWidth: 80, //列宽
            dataType: 10, //默认数据类型
            dataSet: 1, //默认数据设置
            padTop: 2, //上边距
            padBottom: 2, //下边距
            padLeft: 2, //左边距
            padRight: 2, //右边距
            divId: "canvas", //绑定的Div id
            doneFlag: false
        },
        tools = {
            apply: function (fun, param, defaultValue) {
                if ((typeof fun) == "function") {
                    return fun.apply(zt, param ? param : []);
                }
                return defaultValue;
            },
            clone: function (obj) {
                if (obj === null) return null;
                var o = tools.isArray(obj) ? [] : {};
                for (var i in obj) {
                    o[i] = (obj[i] instanceof Date) ? new Date(obj[i].getTime()) : (typeof obj[i] === "object" ? tools.clone(obj[i]) : obj[i]);
                }
                return o;
            },
            eqs: function (str1, str2) {
                return str1.toLowerCase() === str2.toLowerCase();
            },
            isArray: function (arr) {
                return Object.prototype.toString.apply(arr) === "[object Array]";
            },
            isElement: function (o) {
                return (
                    typeof HTMLElement === "object" ? o instanceof HTMLElement : //DOM2
                        o && typeof o === "object" && o !== null && o.nodeType === 1 && typeof o.nodeName === "string"
                );
            }
        },
        spreadBaseTools = {
            // encoder string
            encode: function (str, _ss) {
                let lengthBytes = this.lengthBytesUTF8(str) + 1;
                var stringOnWasmHeap = _ss._malloc(lengthBytes);
                this.stringToUTF8(str, stringOnWasmHeap, lengthBytes + 1, _ss);
                return stringOnWasmHeap;
            },
            decodeStrAndFree: function (visualIndex, _ss) {
                var str = this.UTF8ToString(visualIndex, _ss);
                _ss._free(visualIndex);
                return str;
            },
            UTF8ToString: function (ptr, _ss, maxBytesToRead) {
                return ptr ? this.UTF8ArrayToString(_ss.HEAPU8, ptr, maxBytesToRead, _ss) : ""
            },
            UTF8ArrayToString: function (heapOrArray, idx, maxBytesToRead, _ss) {
                var endIdx = idx + maxBytesToRead;
                var endPtr = idx;
                while (heapOrArray[endPtr] && !(endPtr >= endIdx)) ++endPtr;
                if (endPtr - idx > 16 && heapOrArray.buffer && _ss.UTF8Decoder) {
                    return _ss.UTF8Decoder.decode(heapOrArray.subarray(idx, endPtr))
                }
                var str = "";
                while (idx < endPtr) {
                    var u0 = heapOrArray[idx++];
                    if (!(u0 & 128)) {
                        str += String.fromCharCode(u0);
                        continue
                    }
                    var u1 = heapOrArray[idx++] & 63;
                    if ((u0 & 224) == 192) {
                        str += String.fromCharCode((u0 & 31) << 6 | u1);
                        continue
                    }
                    var u2 = heapOrArray[idx++] & 63;
                    if ((u0 & 240) == 224) {
                        u0 = (u0 & 15) << 12 | u1 << 6 | u2
                    } else {
                        u0 = (u0 & 7) << 18 | u1 << 12 | u2 << 6 | heapOrArray[idx++] & 63
                    }
                    if (u0 < 65536) {
                        str += String.fromCharCode(u0)
                    } else {
                        var ch = u0 - 65536;
                        str += String.fromCharCode(55296 | ch >> 10, 56320 | ch & 1023)
                    }
                }
                return str
            },
            lengthBytesUTF8: function (str) {
                var len = 0;
                for (var i = 0; i < str.length; ++i) {
                    var c = str.charCodeAt(i);
                    if (c <= 127) {
                        len++
                    } else if (c <= 2047) {
                        len += 2
                    } else if (c >= 55296 && c <= 57343) {
                        len += 4;
                        ++i
                    } else {
                        len += 3
                    }
                }
                return len
            },
            stringToUTF8: function (str, outPtr, maxBytesToWrite, _ss) {
                return this.stringToUTF8Array(str, _ss.HEAPU8, outPtr, maxBytesToWrite)
            },
            stringToUTF8Array: function (str, heap, outIdx, maxBytesToWrite) {
                if (!(maxBytesToWrite > 0)) return 0;
                var startIdx = outIdx;
                var endIdx = outIdx + maxBytesToWrite - 1;
                for (var i = 0; i < str.length; ++i) {
                    var u = str.charCodeAt(i);
                    if (u >= 55296 && u <= 57343) {
                        var u1 = str.charCodeAt(++i);
                        u = 65536 + ((u & 1023) << 10) | u1 & 1023
                    }
                    if (u <= 127) {
                        if (outIdx >= endIdx) break;
                        heap[outIdx++] = u
                    } else if (u <= 2047) {
                        if (outIdx + 1 >= endIdx) break;
                        heap[outIdx++] = 192 | u >> 6;
                        heap[outIdx++] = 128 | u & 63
                    } else if (u <= 65535) {
                        if (outIdx + 2 >= endIdx) break;
                        heap[outIdx++] = 224 | u >> 12;
                        heap[outIdx++] = 128 | u >> 6 & 63;
                        heap[outIdx++] = 128 | u & 63
                    } else {
                        if (outIdx + 3 >= endIdx) break;
                        heap[outIdx++] = 240 | u >> 18;
                        heap[outIdx++] = 128 | u >> 12 & 63;
                        heap[outIdx++] = 128 | u >> 6 & 63;
                        heap[outIdx++] = 128 | u & 63
                    }
                }
                heap[outIdx] = 0;
                return outIdx - startIdx
            },
            getRatio: function () { //获取屏幕比例
                var ratio = 0;
                var screen = window.screen;
                var ua = navigator.userAgent.toLowerCase();

                if (window.devicePixelRatio !== undefined) {
                    ratio = window.devicePixelRatio;
                } else if (~ua.indexOf('msie')) {
                    if (screen.deviceXDPI && screen.logicalXDPI) {
                        ratio = screen.deviceXDPI / screen.logicalXDPI;
                    }

                } else if (window.outerWidth !== undefined && window.innerWidth !== undefined) {
                    ratio = window.outerWidth / window.innerWidth;
                }
                return ratio;
            },
            isFullySelected: function (obj) {
                var text = obj.val();
                var start = obj.prop('selectionStart');
                var end = obj.prop('selectionEnd');
                return start === 0 && end === text.length;
            },

            getPartValue: function (obj) {
                var text = obj.val();
                var start = obj.prop('selectionStart');
                var end = obj.prop('selectionEnd');
                let prefix = text.substring(0, start); //首部分
                let suffix = text.substring(end, text.length);
                return {prefix: prefix, suffix: suffix};
            },

            isNoneSelected: function (obj) {
                var text = obj.val();
                var start = obj.prop('selectionStart');
                var end = obj.prop('selectionEnd');
                return start == end;
            }
        },
        //method of init resource
        resource = {
            // load spreadSheet font
            loadFonts: function (sheetObj, fontUrl, fontFamilyId, fontNameUrl) {
                sheetObj._removeAllFonts(); //remove
                if (fontUrl == '') {
                    return;
                }
                if (tools.isArray(fontUrl)) { //load multiple js
                    $.each(fontUrl, function (i, e) {
                        resource.loadFont(sheetObj, e, i, fontFamilyId, fontNameUrl);
                    })
                } else {
                    resource.loadFont(sheetObj, fontUrl, 0, fontFamilyId, fontNameUrl);
                }
            },
            // load spreadSheet font
            loadFont: function (obj, fontUrl, index, fontFamilyId, fontNameUrl) {
                if (obj != null) {
                    fetch(fontUrl, {
                        method: 'GET',
                        responseType: 'arraybuffer'
                    }).then(function (response) {
                        if (response.status == 200) { //首先返回数据
                            return response.arrayBuffer();
                        }
                    }).then(function (ab) { //ab即为返回的数据
                        var data = new Uint8Array(ab);
                        var len = data.length;
                        var buf = obj._malloc(len);
                        obj.HEAPU8.set(data, buf);
                        var res = obj._loadFont(buf, len);

                        //init fontFamily Select
                        if (fontFamilyId != undefined) {
                            if (fontFamilyId != '') {
                                resource.initFontFamily(fontNameUrl, obj, fontFamilyId);
                            }
                        }

                        if (index == 0) {
                            if(example != ''){
                                $.ajax({
                                    url: example, // 文件的URL
                                    type: 'GET', // 请求类型，通常是GET
                                    dataType: 'text', // 期望的返回数据类型，这里是文本
                                    success: function(data) {
                                        //obj._addSpreadSheet(spreadBaseTools.encode('',obj), spreadBaseTools.encode('',obj));
                                        //obj._restoreFromJsonStream(spreadBaseTools.encode(data , obj))
                                        obj._addSpreadSheet(spreadBaseTools.encode('',obj), spreadBaseTools.encode(data , obj));
                                        obj._setTabDisplay(false);
                                        obj._setLogicalZoom(spreadBaseTools.getRatio());
                                        obj._setAddFormBtnHidden(true); //hide form button
                                        obj._setShowPageLine(false);
                                        obj._setShowColRowShadow(false);
                                        obj._setAutoPaperSize(true);
                                      /*  setTimeout(function(){

                                        },1000)*/
                                    },
                                    error: function(jqXHR, textStatus, errorThrown) {
                                        console.error('Error: ' + textStatus, errorThrown);
                                    }

                                });
                            }else{
                                obj._addSpreadSheet(spreadBaseTools.encode('',obj), spreadBaseTools.encode('',obj));
                                obj._setTabDisplay(false);
                                obj._setLogicalZoom(spreadBaseTools.getRatio());
                                obj._setAddFormBtnHidden(true); //hide form button
                                obj._setShowPageLine(false);
                                obj._setShowColRowShadow(false);
                                obj._setAutoPaperSize(true);

                            }


                            let copyAreaHtml = '<input type="text" id="ef_copy_area" style="position: absolute;top: 0;left: 0;width: 1px;height: 1px;border:none;background-color: transparent;">';
                            if ($('#ef_copy_area').length == 0) {
                                document.body.insertAdjacentHTML('beforeend', copyAreaHtml);
                            }

                            _variables.timer = setInterval(function () {
                                if (_variables.jsDoneFlag == _variables.jsLength) {
                                    for (var i = 0; i < _variables.jsLength; i++) {
                                        var ba = _variables.jsContentMap[i];
                                        var data = new Uint8Array(ba);
                                        var len = data.length;
                                        var buf = obj._malloc(len);
                                        obj.HEAPU8.set(data, buf);
                                        var res = obj._loadClacExpr(buf, len);
                                        obj._free(buf);
                                    }
                                   /* obj._setShowExprValue(false);
                                    obj._setShowExprValue(true);*/
                                    window.clearInterval(_variables.timer);
                                }
                            }, 200);
                        }
                    });

                }
            },
            // load spreadSheet javascript
            loadJavascripts: function (url) {
                if (url == '') {
                    return;
                }
                if (tools.isArray(url)) { //load multiple js
                    var num = url.length;
                    _variables.jsLength = num;
                    for (var i = 0; i < num; i++) {
                        var u = url[i];
                        resource.loadJavascript(u, i);
                    }
                }

            },
            // fetch js
            loadJavascript: function (url, index) {
                fetch(url, {
                    method: 'GET',
                    responseType: 'arraybuffer'
                }).then(function (response) {
                    if (response.status == 200) { //首先返回数据
                        return response.arrayBuffer();
                    }
                }).then(function (ab) { //ab即为返回的数据
                    _variables.jsDoneFlag++;
                    _variables.jsContentMap[index] = ab;
                });
            },
            // loadDone
            initFontFamily: function (url, obj, fontFamilyId) {
                $.get(url, function (data) {
                    let fontMapVK = {};
                    var dataJson = JSON.parse(data);
                    for (var key in dataJson) {
                        fontMapVK[dataJson[key]] = key;
                    }
                    obj._setRealFontNameList(spreadBaseTools.encode(data, obj));
                    var ts = obj._getAllFontName();
                    ts = spreadBaseTools.decodeStrAndFree(ts, obj);
                    ts = JSON.parse(ts).family;
                    let fontFamilyObj = $('#' + fontFamilyId);
                    fontFamilyObj.empty();
                    for (var x = ts.length - 1; x >= 0; x--) {
                        var name = ts[x];
                        fontFamilyObj.append('<option value="' + name + '">' + (fontMapVK[name] == undefined ? name : fontMapVK[name]) + '</option>');
                    }
                    fontFamilyObj.val('宋体');
                });
            }
        },
        //method of operate input
        view = {};
    // spreadSheet defined
    $.fn.spreadSheet = {
        _z: {
            spreadTool: spreadBaseTools
        },
        view: {
            createInput: function (obj, inputContainer, input) {
                let container = obj.parent();
                let timestamp = Date.now();
                let containerId = 'textInputDiv_' + timestamp;
                let inputId = 'textInput_' + timestamp;
                let inputHtml = " <div id='" + containerId + "' style='display: none;position: absolute;width:0px;height:0px;border:none;'>\n" +
                    "                      <textarea id='" + inputId + "' autocomplete='off' style='resize: none;overflow:hidden;border:none;z-index: 10000;' wrap='hard'></textarea>\n" +
                    "                    </div>";
                container.append(inputHtml);
                inputContainer = $('#' + containerId); //init inputContainerObj
                input = $('#' + inputId); //init inputObj
                this.bindInput(input);
            },
            bindInput: function (obj, ss) {
                //7.5使用，用来使设计器界面上的input获取焦点，从而可以接收到复制粘贴的事件
                document.addEventListener('copy', (e) => {
                    let id = $(e.target).attr('id');
                    var explorer = navigator.userAgent.toLowerCase();
                    if (explorer.indexOf("chrome") >= 0) {
                        e.preventDefault();
                        e.stopPropagation();
                        let text = document.getSelection().toString();
                        e.clipboardData.setData("text", text);
                    }
                });

                document.addEventListener('cut', (e) => {
                    var explorer = navigator.userAgent.toLowerCase();
                    if (explorer.indexOf("chrome") >= 0) {
                        e.preventDefault();
                        e.stopPropagation();
                        let element = document.activeElement;
                        let cursorPos = $(element)[0].selectionStart; //获取光标的位置
                        let text = document.getSelection().toString(); //获取截取文本长度
                        if (text == "") { //解决多次绑定时，ctrl+x失效的bug

                        } else {
                            let originValue = element.value; //原始位置
                            let prefix = originValue.substring(0, cursorPos); //首部分
                            let end = originValue.substring(cursorPos + text.length, originValue.length);
                            let newValue = prefix + end;
                            e.clipboardData.setData("text", text);

                            element.value = newValue;
                        }

                    }
                });

                document.addEventListener('paste', (e) => {
                    let id = $(e.target).attr('id');

                    var explorer = navigator.userAgent.toLowerCase();
                    if (explorer.indexOf("chrome") >= 0) {
                        e.preventDefault();
                        e.stopPropagation();

                        let targetType = $(e.target).prop('type');
                        let targetId = $(e.target).attr('id');


                        var clipboardData = e.clipboardData;
                        var pastedText = clipboardData.getData('text/plain');
                        // 获取你想要触发paste事件的元素
                        var element = document.activeElement;
                        let cursorPos = $(element)[0].selectionStart; //获取光标的位置
                        let originValue = element.value; //原始值
                        let prefix = originValue.substring(0, cursorPos); //首部分
                        let end = originValue.substring(cursorPos, originValue.length);
                        if (id == obj.attr("id")) {
                            if ($(e.target).width() == 0) {
                                element.value = element.value + pastedText;
                            } else {
                                if (spreadBaseTools.isFullySelected($(e.target))) { //全选状态下粘贴，应该覆盖掉当前值
                                    element.value = pastedText;
                                } else {
                                    if (spreadBaseTools.isNoneSelected($(e.target))) { //未选择文本
                                        element.value = prefix + pastedText + end;
                                    } else { //部分选择文本
                                        let val = spreadBaseTools.getPartValue($(e.target));
                                        element.value = val.prefix + pastedText + val.suffix;
                                    }

                                }
                                ss._setSelCellText(spreadBaseTools.encode(element.value, ss)); //设置当前选择的单元格文本
                            }
                        } else {
                            if (spreadBaseTools.isFullySelected($(e.target))) { //全选状态下粘贴，应该覆盖掉当前值
                                element.value = pastedText;
                            } else {
                                if (spreadBaseTools.isNoneSelected($(e.target))) { //未选择文本
                                    element.value = prefix + pastedText + end;
                                } else { //部分选择文本
                                    let val = spreadBaseTools.getPartValue($(e.target));
                                    element.value = val.prefix + pastedText + val.suffix;
                                }
                            }
                        }
                    }


                });
                let views = this;
                obj.bind('paste', function (event) {
                    obj.focus();
                    let sheetType = ss._currSheetType(); //获取当前sheet页类型
                    if (sheetType == 4) { //form表单,无需判断是否允许编辑
                        setTimeout(function () {
                            views.canvasPaste(copyContent, ss); //调用设计器事件，将文本输入框的值赋给设计器并调用paste事件
                        }, 100)
                    } else {
                        if (ss._isAllowEditCurrCell()) { //单元格允许编辑
                            if (obj.width() == 0) {
                                obj.val('');
                                setTimeout(function () {
                                    let value = obj.val();
                                    views.canvasPaste(value, ss); //调用设计器事件，将文本输入框的值赋给设计器并调用paste事件
                                }, 100)
                            } else {

                            }
                        }
                    }
                });

                //文本输入框修改事件
                obj.bind('keyup', function (event) {
                        let directions = [37, 38, 39, 40, 13];
                        if (!event.ctrlKey) {
                            //keyup不处理复制粘贴剪切，交给keydown去处理
                            //if (event.keyCode != 67 && event.keyCode != 86 && event.keyCode != 88 && event.keyCode != 90) {
                            if ($.inArray(event.keyCode, directions) != -1) { //上下左右方向键
                                //文本输入框不处于编辑状态
                                if (obj.parent().width() == 0) {
                                    //调用canvas方法来实现单元格的上下左右移动
                                    views.canvasKeyPress(5, event, ss);
                                    //移动后，更新文本输入框的值为当前单元格的值
                                    $(this).val(ss._getSelCellText());
                                    //处理上下左右按键到表格底部和最右边时丢失焦点的问题
                                    setTimeout(function () {
                                        let canvas = document.getElementById('canvas');
                                        $(canvas).trigger('click'); //触发单击事件获取编辑框焦点 whj
                                        obj.focus();
                                    }, 100)
                                }
                            } else {
                                if (event.keyCode == 17) { //Control

                                } else {
                                    if (event.keyCode == 46) { // DELETE
                                        ss._removeSelCellData(false, false); //删除单元格内容
                                        ss._removeSelShapePlugin();//删除悬浮元素
                                    } else {
                                        if (event.keyCode != 86) {
                                            let text = $(this).val(); //获取输入框内的值
                                            ss._setSelCellText(spreadBaseTools.encode(text, ss)); //设置当前选择的单元格文本
                                        }

                                    }
                                }
                            }
                            //}
                        } else {

                        }
                    }
                )
                //用keydown来监听复制粘贴剪切事件
                obj.bind('keydown', function (event) {
                        if (!event.ctrlKey) {
                            if (event.keyCode == 86) { // v
                                let text = obj.val(); //获取输入框内的值
                                ss._setSelCellText(spreadBaseTools.encode(text, ss)); //设置当前选择的单元格文本
                            }
                        }
                        if (event.ctrlKey && event.keyCode == 67) { //ctrl + c事件、
                            if ($(this).width() == 0) { //在单元格上复制
                                views.canvasKeyPress(5, event, ss); //调用设计器的keypressevent，来触发clipboardCopyEvt事件，将设计器上获取的值赋给文本输入框
                            } else { //在EFTextInput内复制

                            }
                        } else if (event.ctrlKey && event.keyCode == 86) {//ctrl + v事件
                            //event.stopPropagation(); //阻止冒泡
                        } else if (event.ctrlKey && event.keyCode == 88) {//ctrl + x事件
                            if ($(this).width() == 0) { //在单元格上复制
                                views.canvasKeyPress(5, event, ss); //调用设计器的keypressevent，来触发clipboardCopyEvt事件，将设计器上获取的值赋给文本输入框
                            } else { //在EFTextInput内复制

                            }
                        } else if (event.ctrlKey && event.keyCode == 65) {//ctrl + A事件
                            if (obj.width() == 0) { //input未显示状态下，才执行全选操作，否则默认input的全选操作
                                views.canvasKeyPress(5, event, ss);
                            }
                        } else if (event.ctrlKey && event.keyCode == 46) { //ctrl + delete ，清空内容和样式
                            ss._removeSelCellData(true, true);
                            ss._removeSelShapePlugin();
                        } else if (event.ctrlKey && event.keyCode == 90) { //ctrl + z
                            ss._undo();
                        } else if (event.ctrlKey && event.keyCode == 83) { //ctrl + s
                            event.preventDefault(); //屏蔽浏览器默认事件
                        }
                    }
                )
            },
            canvasPaste: function (content, ss) {
                var str = spreadBaseTools.encode(content, ss);
                ss._copyClipboardDataToSpreadsheet(str);
                ss._paste();
            },
            canvasKeyPress: function (type, event, ss) {
                if (event.keyCode == '86' && event.ctrlKey) {
                    return;
                }
                var data = {
                    type: type
                    , x: (event.clientX || 0)
                    , y: (event.clientY || 0)
                    , ctrlKey: event.ctrlKey
                    , shiftKey: event.shiftKey
                    , key: event.key || ''
                    , keyCode: event.keyCode || 0
                    , button: event.button || 0
                };
                var d = JSON.stringify(data);
                if (type == 3 || type == 4 || type == 5) {
                    ss._keyMouseEvent(spreadBaseTools.encode(d, ss));
                } else {
                    //canvas区域内
                    if (data.y > 0 && ($("canvas").width() - data.x > 16) && ($("canvas").height() - 25 - data.y > 16)) {
                        ss._keyMouseEvent(spreadBaseTools.encode(d, ss));
                        if (e.button == 2) {
                            showMenu(e.clientX, e.clientY);
                        }
                    }
                }
            },
            //hider editor
            hideEditor: function (ss) {
                objs.spreadInputContainer.css({
                    left: 0,
                    top: 0
                }).hide();
                objs.spreadInput.css({
                    left: 0,
                    top: 0
                }).val("").blur().hide();
                ss._setSpreadFocus(); //设计器获取焦点
            },
            showEditor: function (left, top, width, height, ss) {
                objs.spreadInputContainer.attr('width', width).attr('height', height).css({ //在canvas上显示文本编辑框DIV
                    "display": "block",
                    "border": "1px solid #44B4FF",
                    "top": top, //25是设计器tab的高度
                    "left": left,
                    "width": width - 1,
                    "height": height - 1,
                    "z-index": 10000
                }).show();
                //设置文本输入框的大小
                objs.spreadInput.css({ //在canvas上显示文本编辑框DIV
                    "width": width - 1,
                    "height": height - 1,
                    "border": 0,
                    "padding": 0,
                    "z-index": 10000
                }).show();

                view.bindInput(objs.spreadInput, ss)

            }
        },
        // spreadSheet init
        initSheet: function (obj, wasmUrl, fontUrl, jsUrl, sheetSetting, fontFamilyId, fontNameUrl) {
            var spread = this;
            var spreadSheetObj;
            var inputContainer;
            var input;
            let container = obj.parent();
            let timestamp = Date.now();
            let containerId = 'textInputDiv_' + timestamp;
            let inputId = 'textInput_' + timestamp;
            let inputHtml = " <div id='" + containerId + "' style='display: none;position: absolute;width:0px;height:0px;border:none;'>\n" +
                "                    <textarea id='" + inputId + "' autocomplete='off' style='resize: none;overflow:hidden;border:none' wrap='hard'></textarea>\n" +
                "                </div>";
            container.append(inputHtml);

            inputContainer = $('#' + containerId); //init inputContainerObj
            input = $('#' + inputId); //init inputObj

            var setting = tools.clone(_setting); //copy setting
            $.extend(true, setting, sheetSetting); //merge setting
            let canvas = obj[0];
            // qtLoader init
            let qtLoader = QtLoader({
                canvasElements: [canvas],
                showLoader: function (loaderStatus) {
                    canvas.style.display = 'none';
                },
                showError: function (errorText) {
                    canvas.style.display = 'none';
                },
                showExit: function () {
                    canvas.style.display = 'none';
                },
                showCanvas: function () {
                    canvas.style.display = 'block';
                    setting.divId = obj.attr("id");
                    resource.loadFonts(sObj, fontUrl, fontFamilyId, fontNameUrl);
                    resource.loadJavascripts(jsUrl);
                }
            });
            var sObj = qtLoader.loadEmscriptenModule(wasmUrl);
            this.view.bindInput(input, sObj);
            spreadSheetObj = sObj;
            //spread = sheetObj;
            var spreadSheetInstance = {
                parent: spreadSheetObj,
                ss: sObj,
                inputContainer: inputContainer,
                input: input,
                getSpreadSheetObj: function () {
                    return sObj;
                },
                encode: function (str) {
                    let lengthBytes = this.lengthBytesUTF8(str) + 1;
                    var stringOnWasmHeap = this.ss._malloc(lengthBytes);
                    this.stringToUTF8(str, stringOnWasmHeap, lengthBytes + 1);
                    return stringOnWasmHeap;
                },
                lengthBytesUTF8: function (str) {
                    var len = 0;
                    for (var i = 0; i < str.length; ++i) {
                        var c = str.charCodeAt(i);
                        if (c <= 127) {
                            len++
                        } else if (c <= 2047) {
                            len += 2
                        } else if (c >= 55296 && c <= 57343) {
                            len += 4;
                            ++i
                        } else {
                            len += 3
                        }
                    }
                    return len
                },
                stringToUTF8: function (str, outPtr, maxBytesToWrite) {
                    return this.stringToUTF8Array(str, this.ss.HEAPU8, outPtr, maxBytesToWrite)
                },
                decodeStrAndFree: function (visualIndex) {
                    var str = this.UTF8ToString(visualIndex);
                    this.ss._free(visualIndex);
                    return str;
                },
                UTF8ToString: function (ptr, maxBytesToRead) {
                    return ptr ? this.UTF8ArrayToString(this.ss.HEAPU8, ptr, maxBytesToRead) : ""
                },
                UTF8ArrayToString: function (heapOrArray, idx, maxBytesToRead) {
                    var endIdx = idx + maxBytesToRead;
                    var endPtr = idx;
                    while (heapOrArray[endPtr] && !(endPtr >= endIdx)) ++endPtr;
                    if (endPtr - idx > 16 && heapOrArray.buffer && this.ss.UTF8Decoder) {
                        return this.ss.UTF8Decoder.decode(heapOrArray.subarray(idx, endPtr))
                    }
                    var str = "";
                    while (idx < endPtr) {
                        var u0 = heapOrArray[idx++];
                        if (!(u0 & 128)) {
                            str += String.fromCharCode(u0);
                            continue
                        }
                        var u1 = heapOrArray[idx++] & 63;
                        if ((u0 & 224) == 192) {
                            str += String.fromCharCode((u0 & 31) << 6 | u1);
                            continue
                        }
                        var u2 = heapOrArray[idx++] & 63;
                        if ((u0 & 240) == 224) {
                            u0 = (u0 & 15) << 12 | u1 << 6 | u2
                        } else {
                            u0 = (u0 & 7) << 18 | u1 << 12 | u2 << 6 | heapOrArray[idx++] & 63
                        }
                        if (u0 < 65536) {
                            str += String.fromCharCode(u0)
                        } else {
                            var ch = u0 - 65536;
                            str += String.fromCharCode(55296 | ch >> 10, 56320 | ch & 1023)
                        }
                    }
                    return str
                },
                stringToUTF8Array: function (str, heap, outIdx, maxBytesToWrite) {
                    if (!(maxBytesToWrite > 0)) return 0;
                    var startIdx = outIdx;
                    var endIdx = outIdx + maxBytesToWrite - 1;
                    for (var i = 0; i < str.length; ++i) {
                        var u = str.charCodeAt(i);
                        if (u >= 55296 && u <= 57343) {
                            var u1 = str.charCodeAt(++i);
                            u = 65536 + ((u & 1023) << 10) | u1 & 1023
                        }
                        if (u <= 127) {
                            if (outIdx >= endIdx) break;
                            heap[outIdx++] = u
                        } else if (u <= 2047) {
                            if (outIdx + 1 >= endIdx) break;
                            heap[outIdx++] = 192 | u >> 6;
                            heap[outIdx++] = 128 | u & 63
                        } else if (u <= 65535) {
                            if (outIdx + 2 >= endIdx) break;
                            heap[outIdx++] = 224 | u >> 12;
                            heap[outIdx++] = 128 | u >> 6 & 63;
                            heap[outIdx++] = 128 | u & 63
                        } else {
                            if (outIdx + 3 >= endIdx) break;
                            heap[outIdx++] = 240 | u >> 18;
                            heap[outIdx++] = 128 | u >> 12 & 63;
                            heap[outIdx++] = 128 | u >> 6 & 63;
                            heap[outIdx++] = 128 | u & 63
                        }
                    }
                    heap[outIdx] = 0;
                    return outIdx - startIdx
                },
                setLogicalZoom: function (height) {
                    this.ss._setLogicalZoom(this.getRatio());
                },
                setEventName: function (name) {
                    let instance = this;
                    let obj = this.ss;
                    let timer = setInterval(function () {
                        if (obj.qtFontDpi != undefined) {
                            obj._setJSObjectName(instance.encode(name));
                            window.clearInterval(timer);
                        }

                    }, 1000)

                },
                hideEditor: function () {
                    this.inputContainer.css({
                        left: 0,
                        top: 0
                    }).hide();
                    this.input.css({
                        left: 0,
                        top: 0
                    }).val("").blur().hide();
                },
                showEditor: function (left, top, width, height, curText) {
                    this.inputContainer.attr('width', width).attr('height', height).css({ //在canvas上显示文本编辑框DIV
                        "display": "block",
                        "border": "1px solid #44B4FF",
                        "top": top, //25是设计器tab的高度
                        "left": left,
                        "width": width - 1,
                        "height": height - 1,
                        "z-index": 10000
                    }).show();
                    let ii = this.input;
                    //设置文本输入框的大小
                    ii.css({ //在canvas上显示文本编辑框DIV
                        "width": width - 1,
                        "height": height - 1,
                        "border": 0,
                        "padding": 0,
                        "z-index": 10000
                    }).show();
                    let text = this.getSelCellText(); //获取当前单元格文本
                    if (curText == undefined) {
                        ii.val(text); //设置文本值
                        //获取编辑焦点
                        setTimeout(function () {
                            ii.focus().val("").val(text); //光标跳到末尾
                        }, 50)
                    } else {
                        //获取编辑焦点
                        setTimeout(function () {
                            ii.focus().val("").val(text + curText); //光标跳到末尾
                            //bindInput();
                        }, 50)
                    }

                },
                mainPaste: function () {

                },
                inputFocus: function () { //input focus
                    this.input.focus();
                },
                getRatio: function () {
                    return spreadBaseTools.getRatio();
                },
                isAllowEditCurrCell: function () {
                    return this.ss._isAllowEditCurrCell();
                },
                getSelCellRect: function () {
                    let str = this.ss._getSelCellRect();
                    str = this.decodeStrAndFree(str);
                    if (str != '') {
                        selRect = JSON.parse(str);
                        return selRect;
                    }
                },
                getSelBeginCell: function () {
                    return this.ss._getSelBeginCell();
                },
                getSelCellText: function () {
                    let sval = this.ss._getSelCellText();
                    var text = this.decodeStrAndFree(sval);
                    return text;
                },
                currSheetType: function () {
                    return this.ss._currSheetType();
                },
                setRowHeight: function (val) {
                    return this.ss._setRowHeight(val);
                },
                getRowHeight: function () {
                    return this.ss._getRowHeight();
                },
                getCurrentSheet: function () {
                    return this.ss._getCurrentSheet();
                },
                insertRow: function () {
                    let index = this.getCurrentSheet();
                    this.ss._insertSheetRow(index);
                },
                appendRow: function () {
                    let index = this.getCurrentSheet();
                    this.ss._appendSheetRow(index);
                },
                deleteRow: function () {
                    let index = this.getCurrentSheet();
                    this.ss._removeSheetRow(index);
                },
                //column method
                setColumnWidth: function (val) {
                    return ss._setColumnWidth(val);
                },
                getColumnWidth: function () {
                    return ss._getColumnWidth();
                },
                insertColumn: function () {
                    let index = this.getCurrentSheet();
                    this.ss._insertSheetColumn(index);
                },
                appendColumn: function () {
                    let index = this.getCurrentSheet();
                    this.ss._appendSheetColumn(index);
                },
                deleteColumn: function () {
                    let index = this.getCurrentSheet();
                    this.ss._removeSheetColumn(index);
                },
                /**
                 *
                 * cell method
                 *
                 * **/
                setSelCellText: function (val) {
                    this.ss._setSelCellText(this.encode(val));
                },
                moveSelFrameToCell: function (val) {
                    return this.ss._moveSelFrameToCell(this.encode(val))
                },
                getCellPos: function () {
                    let pos = this.decodeStrAndFree(this.ss._getSelBeginCell());
                    let posJ = JSON.parse(pos);
                    let cellPos = this.ss._cellPos2Char(posJ.x, posJ.y);
                    return this.decodeStrAndFree(cellPos);
                },
                //获取当前单元格
                getSelCellRect: function () {
                    let str = this.ss._getSelCellRect();
                    str = this.decodeStrAndFree(str);
                    if (str != '') {
                        selRect = JSON.parse(str);
                        return selRect;
                    }
                },
                //设置单元格类型，type=9时，插入斜线
                setSelCellsType: function (type) {
                    return this.ss._setSelCellsType(type);
                },
                //设置单元格样式
                setSelCellsStyle: function (type) {
                    return this.ss._setSelCellsStyle(type);
                },
                //1.获取单元格样式,return int (HCellStyle)
                getSelCellsStyle: function () {
                    return this.ss._getSelCellsStyle();
                },
                //设置二维码类型
                setSelCellsBarCodeType: function (type) {
                    return this.ss._setSelCellsBarCodeType(type);
                },
                getSelCellsBarCodeHideText: function () {
                    return this.ss._getSelCellsBarCodeHideText();
                },
                //2.获取二维码类型 return int (BarCodeType)
                getSelCellsBarCodeType: function () {
                    return this.ss._getSelCellsBarCodeType();
                },
                //当前单元格是否是插件
                isSelCellPluginInfo: function () {
                    return this.ss._isSelCellPluginInfo();
                },
                setCellPluginDefaultImage: function (col, row, base64) {
                    return this.ss._setCellPluginDefaultImage(col, row, this.encode(base64));
                },
                setSelCellPluginInfo: function (option) {
                    return this.ss._setSelCellPluginInfo(this.encode(option));
                },
                getPluginInfo: function () {
                    return this.ss._pluginInfo();
                },
                //获取插件类型
                getPluginType: function () {
                    return this.decodeStrAndFree(this.ss._pluginType());
                },
                //获取插件类型
                getPluginName: function () {
                    return this.decodeStrAndFree(this.ss._pluginName());
                },
                //当前单元格是否允许编辑
                isAllowEditCurrCell: function () {
                    return this.ss._isAllowEditCurrCell();
                },
                //当前单元格位置
                getSelBeginCell: function () {
                    return this.ss._getSelBeginCell();
                },
                //获取当前单元格文本
                getSelCellText: function () {
                    let sval = this.ss._getSelCellText();
                    var text = this.decodeStrAndFree(sval);
                    return text;
                },
                //设置单元格文本
                setSelCellText: function (val) {
                    this.ss._setSelCellText(this.encode(val));
                },
                //删除当前单元格内容
                removeSelCellData: function () {
                    return this.ss._removeSelCellData(false, false);
                },
                //设置单元格注释
                setSelCellComment: function (val) {
                    return this.ss._setSelCellComment(this.encode(val));
                },
                //设置单元格注释
                getSelCellComment: function () {
                    return this.decodeStrAndFree(this.ss._getSelCellComment());
                },
                //单元格注释是否隐藏
                getSelCellCommentHide: function () {
                    return this.ss._getSelCellCommentHide();
                },
                //设置单元格注释是否隐藏
                setSelCellCommentHide: function (val) {
                    return this.ss._setSelCellCommentHide(val);
                },
                setMergeExpandCellStr: function (val) {
                    return this.ss._setSelCellMergeExpandCellStr(val);
                },
                getMergeExpandCellStr: function () {
                    return this.ss._isMergeExpandCellStr();
                },
                //删除单元格内容和样式
                removeSelCell: function () {
                    return this.ss._removeSelCellData(true, true);
                },
                //设置当前单元格字体
                setSelCellFontFamily: function (val) {
                    return this.ss._setSelCellFontFamily(this.encode(val));
                },
                //设置当前单元格字体大小
                setSelCellFontPointSize: function (val) {
                    return this.ss._setSelCellFontPointSize(val);
                },
                //取消所有边框样式
                cancelSelCellLineStyle: function () {
                    this.ss._cancelSelCellLineStyle();
                },
                //设置单元格边框样式
                setSelCellLineStyle: function (sideType, penStyle, widthType, color) {
                    return this.ss._setSelCellLineStyle(sideType, penStyle, widthType, this.encode(color));
                },
                setSelCellLineStyleByPenStyle: function (sideType, penStyle) {
                    this.ss._setSelCellLineStyleByPenStyle(sideType, penStyle);
                },
                //获取单元格边框样式
                getSelCellLineStyle: function (type) {
                    return this.ss._getSelCellLineStyle(type);
                },
                //设置单元格背景颜色
                setSelCellBKColor: function (color) {
                    this.ss._setSelCellBKColor(this.encode(color));
                },
                //设置单元格水平方向
                setSelCellAlignH: function (val) {
                    return this.ss._setSelCellAlignH(val);
                },
                //获取单元格水平方向
                getSelCellAlignH: function () {
                    return this.ss._getSelCellAlignH();
                },
                //设置单元格垂直方向
                setSelCellAlignV: function (align) {
                    return this.ss._setSelCellAlignV(align);
                },
                getSelCellAlignV: function () {
                    return this.ss._getSelCellAlignV();
                },
                //获取单元格背景颜色
                getSelCellBKColor: function () {
                    return this.decodeStrAndFree(this.ss._getSelCellBKColor());
                },
                //设置单元格字体颜色
                setSelCellFontColor: function (color) {
                    return this.ss._setSelCellFontColor(this.encode(color));
                },
                getSelCellFont: function () {
                    return this.decodeStrAndFree(this.ss._getSelCellFont());
                },
                //获取单元格字体颜色
                getSelCellFontColor: function () {
                    return this.decodeStrAndFree(this.ss._getSelCellFontColor());
                },
                //设置斜体
                setSelCellFontItalic: function (val) {
                    return this.ss._setSelCellFontItalic(val);
                },
                //设置weight
                setSelCellFontWeight: function (val) {
                    return this.ss._setSelCellFontWeight(val);
                },
                //设置字体粗体
                setSelCellFontBold: function (val) {
                    return this.ss._setSelCellFontBold(val);
                },
                //设置字体粗体
                setSelCellFontUnderline: function (val) {
                    return this.ss._setSelCellFontUnderline(val);
                },
                //设置单元格自适应高度
                setSelCellAdaptTextHeight: function (val) {
                    return this.ss._setSelCellAdaptTextHeight(val);
                },
                //单元格是否自适应高度
                isSelCellAdaptTextHeight: function () {
                    return this.ss._isSelCellAdaptTextHeight();
                },
                //获取格式化文本
                getSelCellFormatText: function () {
                    return this.decodeStrAndFree(this.ss._getSelCellFormatText());
                },
                //设置单元格背景图片
                setSelCellBkPic: function (val) {
                    this.ss._setSelCellBkPic(this.encode(val));
                },
                //移除背景图片
                removeSelCellBkPic: function () {
                    this.ss._removeSelCellBkPic();
                },
                /**
                 * 数据集相关属性
                 * */
                //获取单元格类型
                getSelCellsType: function () {
                    return this.ss._getSelCellsType();
                },
                //1.获取单元格样式,return int (HCellStyle)
                getSelCellsStyle: function () {
                    return this.ss._getSelCellsStyle();
                },
                //2.判断是否是表结构字段 param 单元格位置X，Y
                isFieldCell: function (x, y) {
                    return this.ss._isFieldCell(x, y);
                },
                //获取当前单元格的数据集名称
                getDSName: function () {
                    var val = this.ss._getDSName();
                    return this.decodeStrAndFree(val);
                },
                //设置当前单元格数据集名称
                setDSName: function (val) {
                    return this.ss._setSelCellDSName(this.encode(val));
                },
                //获取当前单元格的字段名
                getFieldName: function () {
                    var val = this.ss._getFieldName();
                    return this.decodeStrAndFree(val);
                },
                setFieldName: function (val) {
                    return this.ss._setSelCellFieldName(this.encode(val));
                },
                //获取单元格数据类型
                getFieldDataType: function () {
                    return this.ss._getFieldDataType();
                },
                setFieldDataType: function (val) {
                    return this.ss._setSelCellFieldDataType(val);
                },
                //获取数据设置
                getDataAttribution: function () {
                    return this.ss._getDataAttribution();
                },
                setDataAttribution: function (val) {
                    return this.ss._setSelCellDataAttribution(val);
                },
                //获取汇总方式
                getDataStatisticsType: function () {
                    return this.ss._getDataStatisticsType();
                },
                //设置汇总方式
                setDataStatisticsType: function (val) {
                    return this.ss._setSelCellDataStatisticsType(val);
                },
                //获取过滤表达式
                getFilterExpr: function () {
                    var val = this.ss._getFilterExpr();
                    return this.decodeStrAndFree(val);
                },
                //设置最小行数
                setMinRecordCount: function (count) {
                    return this.ss._setSelCellMinRecordCount(count);
                },
                //获取最小显示行数
                getMinRecordCount: function () {
                    return this.ss._getMinRecordCount();
                },
                //获取数据字典ds名称
                getDataDictDsName: function () {
                    var val = this.ss._getDataDictDsName();
                    return this.decodeStrAndFree(val);
                },
                setDataDictDsName: function (val) {
                    return this.ss._setDataDictDsName(this.encode(val));
                },
                setDataDict: function () {
                    return this.ss._setDataDict(this.encode(''), this.encode(''), this.encode(''));
                },
                //获取数据字典映射列
                getActualFieldName: function () {
                    var val = this.ss._getActualFieldName();
                    return this.decodeStrAndFree(val);
                },
                //设置数据字典映射列
                setActualFieldName: function (val) {
                    return this.ss._setActualFieldName(this.encode(val));
                },
                //获取数据字典显示列
                getShowFieldName: function () {
                    var val = this.ss._getShowFieldName();
                    return this.decodeStrAndFree(val);
                },
                //设置数据字典显示列
                setShowFieldName: function (val) {
                    return this.ss._setShowFieldName(this.encode(val));
                },
                //是否字段分页
                isFieldPaging: function () {
                    return this.ss._isFieldPaging();
                },
                setFieldPaging: function (flag) {
                    return this.ss._setSelCellFieldPaging(flag);
                },
                //是否行后分页
                isPagingAfterRow: function () {
                    return this.ss._isPagingAfterRow();
                },
                setPagingAfterRow: function (flag) {
                    return this.ss._setSelCellPagingAtferRow(flag);
                },
                //是否每页补齐行
                isCompleteRowForEveryPage: function () {
                    return this.ss._isCompleteRowForEveryPage();
                },
                setCompleteRowForEveryPage: function (flag) {
                    return this.ss._setSelCellCompleteRowForEveryPage(flag);
                },
                //是否尾页补齐行
                isCompleteRowForLastPage: function () {
                    return this.ss._isCompleteRowForLastPage();
                },
                setCompleteRowForLastPage: function (flag) {
                    return this.ss._setSelCellCompleteRowForLastPage(flag);
                },
                //是否图片字段
                isImageField: function () {
                    return this.ss._isImageField();
                },
                setImageField: function (flag) {
                    return this.ss._setImageField(flag);
                },
                //图片路径
                imagePathStr: function () {
                    var val = this.ss._imagePathStr();
                    return this.decodeStrAndFree(val);
                },
                setImagePathStr: function (str) {
                    return this.ss._setImagePathStr(this.encode(str));
                },
                //是否开启分栏
                setSelCellExpandByIsMutliColumn: function (flag) {
                    return this.ss._setSelCellExpandByIsMutliColumn(flag);
                },
                isExpandByMutliColumn: function () {
                    return this.decodeStrAndFree(this.ss._isExpandByMutliColumn());
                },
                //分栏数量
                setSelCellExpandByColCount: function (flag) {
                    return this.ss._setSelCellExpandByColCount(flag);
                },
                //先行后列
                setSelCellExpandByIsColumnFirst: function (flag) {
                    return this.ss._setSelCellExpandByIsColumnFirst(flag);
                },
                //开启折叠
                setRetractSubCellByRetractSubCell: function (flag) {
                    return this.ss._setRetractSubCellByRetractSubCell(flag);
                },
                //是否折叠
                isRetractSubCell: function () {
                    return this.ss._isRetractSubCell();
                },
                //显示最后行
                setRetractSubCellByShowLastSubCell: function (flag) {
                    return this.ss._setRetractSubCellByShowLastSubCell(flag);
                },
                //是否显示最后行
                isShowLastSubCell: function () {
                    return this.ss._isShowLastSubCell();
                },
                //强制允许编辑
                setSelCellForceAllowedEdit: function (flag) {
                    return this.ss._setSelCellForceAllowedEdit(flag);
                },
                isSelCellForceAllowedEdit: function () {
                    return this.ss._isSelCellForceAllowedEdit()
                },
                //拆分并拷贝
                setSelCellDemergeAndCopyCell: function (flag) {
                    return this.ss._setSelCellDemergeAndCopyCell(flag);
                },
                isDemergeAndCopyCell: function () {
                    return this.ss._isDemergeAndCopyCell();
                },
                setSelCellDataAttrListToGroup: function (flag) {
                    return this.ss._setSelCellDataAttrListToGroup(flag);
                },
                getCellDataAttrListToGroup: function () {
                    return this.ss._getCellDataAttrListToGroup();
                },
                //扩展方向
                setExpandOri: function (type) {
                    return this.ss._setExpandOri(type);
                },
                getExpandOri: function () {
                    return this.ss._getExpandOri();
                },
                setLeftParentCellFrame: function () {
                    return this.ss._setLeftParentCellFrame();
                },
                //左父格类型
                getLeftParentCellType: function () {
                    return this.ss._getLeftParentCellType();
                },
                setLeftParentCellType: function (val) {
                    return this.ss._setLeftParentCellType(val);
                },
                //左父格
                getLeftParentCell: function () {
                    var val = this.ss._leftParentCell();
                    return this.decodeStrAndFree(val);
                },
                setLeftParentCell: function (x, y) {
                    var xx = this.cellX2Int(x);
                    return this.ss._setLeftParentCell(xx, y);
                },
                //上父格类型
                getTopParentCellType: function () {
                    return this.ss._getTopParentCellType();
                },
                setTopParentCellType: function (val) {
                    return this.ss._setTopParentCellType(val);
                },
                //上父格单元格
                getTopParentCell: function () {
                    var val = this.ss._topParentCell();
                    return this.decodeStrAndFree(val);
                },
                setTopParentCell: function (x, y) {
                    var xx = this.cellX2Int(x);
                    return this.ss._setTopParentCell(xx, y);
                },
                setTopParentCellFrame: function () {
                    return this.ss._setTopParentCellFrame();
                },
                //排序类型
                setOrderType: function (type) {
                    return this.ss._setOrderType(type);
                },
                getOrderType: function () {
                    return this.ss._getOrderType();
                },
                //横向伸展
                setSelCellExtendH: function (val) {
                    return this.ss._setSelCellExtendH(val);
                },
                isSelCellExtendH: function () {
                    return this.ss._isSelCellExtendH();
                },
                setSelCellExtendV: function (val) {
                    return this.ss._setSelCellExtendV(val);
                },
                //纵向伸展
                isSelCellExtendV: function () {
                    return this.ss._isSelCellExtendV();
                },
                //设置选中单元格是否显示0值
                setSelCellShowZero: function (flag) {
                    return this.ss._setSelCellShowZero(flag);
                },
                isSelCellShowZero: function () {
                    return this.ss._isSelCellShowZero();
                },
                //空值显示值
                setSelCellNullConvertStr: function (val) {
                    return this.ss._setSelCellNullConvertStr(this.encode(val));
                },
                getSelCellNullConvertStr: function () {
                    return this.decodeStrAndFree(this.ss._getSelCellNullConvertStr());
                },
                //设置选中单元格是否隐藏
                setSelCellHided: function (flag) {
                    this.ss._setSelCellHided(flag);
                },
                isSelCellHided: function () {
                    return this.ss._isSelCellHided();
                },
                //行间距
                setSelCellLineSpacing: function (val) {
                    this.ss._setSelCellLineSpacing(val);
                },
                getSelCellLineSpacing: function () {
                    return this.ss._getSelCellLineSpacing();
                },
                //字符间距
                setSelCellLetterSpacing: function (val) {
                    return this.ss._setSelCellLetterSpacing(val);
                },
                getSelCellLetterSpacing: function () {
                    return this.ss._getSelCellLetterSpacing();
                },
                //段落空格数
                setSelCellParagraphSpaceCount: function (val) {
                    return this.ss._setSelCellParagraphSpaceCount(val);
                },
                getSelCellParagraphSpaceCount: function () {
                    return this.ss._getSelCellParagraphSpaceCount();
                },
                //文本上边距
                setSelCellTopMargin: function (val) {
                    return this.ss._setSelCellTopMargin(val);
                },
                getSelCellTopMargin: function () {
                    return this.ss._getSelCellTopMargin();
                },
                //文本下边距
                setSelCellBottomMargin: function (val) {
                    return this.ss._setSelCellBottomMargin(val);
                },
                getSelCellBottomMargin: function () {
                    return this.ss._getSelCellBottomMargin();
                },
                //文本左边距
                setSelCellLeftMargin: function (val) {
                    return this.ss._setSelCellLeftMargin(val);
                },
                getSelCellLeftMargin: function () {
                    return this.ss._getSelCellLeftMargin();
                },
                //文本右边距
                setSelCellRightMargin: function (val) {
                    return this.ss._setSelCellRightMargin(val);
                },
                getSelCellRightMargin: function () {
                    return this.ss._getSelCellRightMargin();
                },
                //设置选中单元格图片水平对齐方式
                setSelCellImageAlignX: function (val) {
                    return this.ss._setSelCellImageAlignX(val);
                },
                getSelCellImageAlignX: function () {
                    return this.ss._getSelCellImageAlignX();
                },
                //设置选中单元格图片垂直对齐方式
                setSelCellImageAlignY: function (val) {
                    return this.ss._setSelCellImageAlignY(val);
                },
                getSelCellImageAlignY: function () {
                    return this.ss._getSelCellImageAlignY();
                },
                //设置选中单元格是否缩放图片
                setSelCellAdaptImageSize: function (val) {
                    return this.ss._setSelCellAdaptImageSize(val);
                },
                isSelCellAdaptImageSize: function () {
                    return this.ss._isSelCellAdaptImageSize();
                },
                //设置选中单元格图片是否保持原比例
                setSelCellRawImageScale: function (val) {
                    return this.ss._setSelCellRawImageScale(val);
                },
                isSelCellRawImageScale: function (val) {
                    return this.ss._isSelCellRawImageScale(val);
                },
                //控件属性
                setControlInfo: function (str) {
                    return this.ss._setSelCellControlInfo(this.encode(str));
                },
                getControlInfo: function () {
                    let val = this.ss._controlInfo();
                    return this.decodeStrAndFree(val);
                },
                setRegionInfo: function (str) {
                    return this.ss._setRegionInfo(this.encode(str));
                },
                getRegionInfo: function () {
                    let val = this.ss._getRegionInfo();
                    return this.decodeStrAndFree(val);
                },
                //新的区域联动属性
                setRepaintRegions: function (str) {
                    return this.ss._setRepaintRegions(this.encode(str));
                },
                //获取新的区域联动属性
                getRepaintRegions: function () {
                    return this.decodeStrAndFree(this.ss._repaintRegions());
                },
                //获取单元格过滤表达式
                getFilterExpr: function () {
                    return this.decodeStrAndFree(this.ss._getFilterExpr());
                },
                //设置单元格过滤表达式
                setFilterExpr: function (val) {
                    return this.ss._setFilterExpr(this.encode(val));
                },
                //将父格作为过滤条件
                getFilterDependentParent: function () {
                    return this.ss._getFilterDependentParent();
                },
                //设置超链接
                setSelCellHyperlink: function (str) {
                    return this.ss._setSelCellHyperlink(this.encode(str));
                },
                //获取选中单元格超级链接
                getSelCellHyperlink: function () {
                    var val = this.ss._getSelCellHyperlink();
                    return this.decodeStrAndFree(val);
                },
                setSelCellAdaptTextWidth: function (flag) {
                    return this.ss._setSelCellAdaptTextWidth(flag);
                },
                isSelCellAdaptTextWidth: function () {
                    return this.ss._isSelCellAdaptTextWidth();
                },
                isShowExprValue: function () {
                    return this.ss._isShowExprValue();
                },
                setShowExprValue: function (flag) {
                    return this.ss._setShowExprValue(flag);
                },
                isShowColRowShadow: function(){
                    return this.ss._isShowColRowShadow();
                },
                setShowColRowShadow: function(flag){
                    return this.ss._setShowColRowShadow(flag);
                },
                //获取模板内容
                getContent: function () {
                    let sval = this.ss._saveAsToStream();
                    return this.decodeStrAndFree(sval);
                },
                /*
                * 设计器操作
                * **/
                copy: function () {
                    return this.ss._copy();
                },
                copyToSheet: function (str) {
                    return this.ss._copyClipboardDataToSpreadsheet(str);
                },
                cut: function () {
                    return this.ss._cut();
                },
                paste: function () {
                    return this.ss._paste();
                },
                undo: function () {
                    return this.ss._undo();
                },
                comeback: function () {
                    return this.ss._comeback();
                },
                formatBrush: function () {
                    return this.ss._setFormatBrushFrame();
                },
                mergeCell: function () {
                    return this.ss._mergeSelCells();
                },
                splitCell: function () {
                    return this.ss._unmergeSelCells();
                },
                insertRow: function () {
                    let index = this.ss._getCurrentSheet();
                    this.ss._insertSheetRow(index);
                },
                appendRow: function () {
                    let index = this.ss._getCurrentSheet();
                    this.ss._appendSheetRow(index);
                },
                deleteRow: function () {
                    let index = this.ss._getCurrentSheet();
                    this.ss._removeSheetRow(index);
                },
                insertColumn: function () {
                    let index = this.ss._getCurrentSheet();
                    this.ss._insertSheetColumn(index);
                },
                appendColumn: function () {
                    let index = this.ss._getCurrentSheet();
                    this.ss._appendSheetColumn(index);
                },
                deleteColumn: function () {
                    let index = this.ss._getCurrentSheet();
                    this.ss._removeSheetColumn(index);
                },
                restoreFromJsonStream: function(str){
                    this.ss._restoreFromJsonStream(this.encode(str));
                },
                setFixedColRow: function (a, b, c, d) {
                    return this.ss._setFixedColRow(a, b, c, d);
                },
                getFixedColRow: function () {
                    let val = this.ss._getFixedColRow();
                    return this.decodeStrAndFree(val);
                },
                setSelDesignTableRegion: function () {
                    return this.ss._setSelDesignTableRegion();
                },
                setSelTableRegion: function(){
                    return this.ss._setSelTableRegion();
                },
                removeSelDesignTableRegion: function () {
                    return this.ss._removeSelDesignTableRegion();
                },
                removeSelTableRegion: function(){
                    return this.ss._removeSelTableRegion();
                },
                cancelOperateState: function(){
                    return this.ss._cancelOperateState();
                },
                filterTableRegionText: function(val){
                    this.ss._filterTableRegionText(this.encode(val))
                },
                orderTableRegion: function(info , order){
                    this.ss._orderTableRegion(this.encode(info) , order);
                },
                filterTableRegionTextByExpr: function(info){
                    this.ss._filterTableRegionTextByExpr(this.encode(info));
                },
                restoreShowAllRowsInTableRegion: function(info){
                    this.ss._restoreShowAllRowsInTableRegion(this.encode(info));
                },
                setShowPageLine: function(flag){
                    this.ss._setShowPageLine(flag);
                },
                importXlsxStream: function (content) {
                    this.ss._importXlsxStream(this.encode(content), 1);
                },
                removeCurrentSheet: function(){
                    this.ss._removeCurrentSheet();
                },
                setCurrentSheetName: function(val){
                    this.ss._setCurrentSheetName(this.encode(val));
                },
                getCurrentSheetName: function(){
                    return this.decodeStrAndFree(this.ss._getCurrentSheetName());
                },
                appendSheet: function(){
                    this.ss._appendSheet()
                },
                cloneSheetByName: function(val){
                    this.ss._cloneSheetByName(this.encode(val));
                }
            };
            return spreadSheetInstance;
        }
    };

})(jQuery);



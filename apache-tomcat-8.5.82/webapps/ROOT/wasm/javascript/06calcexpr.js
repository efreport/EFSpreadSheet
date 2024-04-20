/**
 * 多维数组转一维，主要用在参数转化
 */
function arrConversion (newArr,...arr) {
  for (var i = 0; i < arr.length; i++) {
    if (Array.isArray(arr[i])) {
      arrConversion(newArr,...arr[i])
    } else {
      newArr.push(arr[i])
    }
  }
}

/**
 * 对数进行上舍入
 */
function CEIL(x){
	return Math.ceil(x);
}

/**
 * 返回 e 的指数
 */
function EXP(x){
	return Math.exp(x);
}

/**
 * 返回 x 的 y 次幂
 */
function POW(x,y){
	return Math.pow(x,y);
}

/**
 * 返回 0 ~ 1 之间的随机数
 */
function RANDOM(){
	return Math.random();
}

/**
 * 返回数字中的最大值
 */
function MAXNUM(){
	return Math.max.apply(null,arguments);
}

/**
 * 返回数字中的最小值
 */
function MINNUM(){
	return Math.min.apply(null,arguments);
}

/**
 * 返回链接的字符串
 */
function CONCAT(){
	var res = new String;
	for(var i=0;i< arguments.length; ++i)
	{
		res += arguments[i];
	}
	return res;
}

/**
 * 返回字符串所在位置索引
 */
function INDEXOF(value,searchvalue,fromindex){
	var string = new String(value);
	return string.indexOf(searchvalue,fromindex);
}

/**
 * 返回一个指定的字符串值最后出现的位置，在一个字符串中的指定位置从后向前搜索。
 */
function LASTINDEXOF(value,searchvalue,fromindex){
	var string = new String(value);
	return string.lastIndexOf(searchvalue,fromindex);
}

/**
 * 在字符串内检索指定的值，或找到一个或多个正则表达式的匹配。
 */
function Match(value,searchvalue){
	var string = new String(value);
	return string.match(searchvalue);
}


/**
 * 提取字符串的某个部分，并以新的字符串返回被提取的部分。
 */
function SLICE(value,start,end){
	var string = new String(value); 
	return string.slice(start,end);
}

/**
 * 把一个字符串分割成字符串数组。
 */
function SPLIT(value,separator,howmany){
	var string = new String(value); 
	return string.split(separator,howmany);
}

/**
 * 在字符串中抽取从 start 下标开始的指定数目的字符。
 */
function SUBSTR(value,start,length){
	var string = new String(value); 
	return string.substr(start,length);
}

/**
 * 提取字符串中介于两个指定下标之间的字符。
 */
function SUBSTRING(value,start,stop){
	var string = new String(value); 
	return string.substring(start,stop);
}

/**
 * 把字符串转换为小写。
 */
function TOLOCALELOWERCASE(value){
	var string = new String(value); 
	return string.toLocaleLowerCase();
}

/**
 * 把字符串转换为大写。
 */
function TOLOCALEUPPERCASE(value){
	var string = new String(value); 
	return string.toLocaleUpperCase();
}

/**
 * 把对象的值转换为数字。
 */
function NUMBER(object){
	return Number(object);
}

/**
 * 把对象的值转换为字符串。
 */
function STRING(object){
	return String(object);
}

/**
 * 返回当天的日期和时间。
 */
function date(){
	var d = new Date();
	return d;
}

/**
 * 返回从 1970 年 1 月 1 日至今的毫秒数。
 */
function GETTIME(){
	var d = new Date();
	return d.getTime();
}

function DAYSDELTA(DateOne,DateTwo){
	var OneMonth = DateOne.substring(5,DateOne.lastIndexOf ('-'));  
    var OneDay = DateOne.substring(DateOne.length,DateOne.lastIndexOf ('-')+1);  
    var OneYear = DateOne.substring(0,DateOne.indexOf ('-'));  
  
    var TwoMonth = DateTwo.substring(5,DateTwo.lastIndexOf ('-'));  
    var TwoDay = DateTwo.substring(DateTwo.length,DateTwo.lastIndexOf ('-')+1);  
    var TwoYear = DateTwo.substring(0,DateTwo.indexOf ('-'));  
  
    var cha=((Date.parse(OneMonth+'/'+OneDay+'/'+OneYear)- Date.parse(TwoMonth+'/'+TwoDay+'/'+TwoYear))/86400000);   
    return Math.abs(cha);  
}

function MONTHSDELTA(DateOne,DateTwo){
    DateOne = DateOne.split('-');
    DateOne = parseInt(DateOne[0]) * 12 + parseInt(DateOne[1]);
	
    DateTwo = DateTwo.split('-');
    DateTwo = parseInt(DateTwo[0]) * 12 + parseInt(DateTwo[1]);
	
    var m = Math.abs(DateOne - DateTwo);
	
    return m;
}

function WEEKSDELTA(DateOne,DateTwo){
	var d = DAYSDELTA(DateOne,DateTwo)
	return parseInt(d/7);
}

function time(date){
	var d = new Date(date);
	var sign1 = "-";
	var sign2 = ":";
	var year = d.getFullYear();
	var month = d.getMonth() + 1; 
	var day  = d.getDate();
	var hour = d.getHours();
	var minutes = d.getMinutes();
	var seconds = d.getSeconds();

	if (month >= 1 && month <= 9) 
	{
		month = "0" + month;
	}
	if (day >= 0 && day <= 9) 
	{
		day = "0" + day;
	}
	if (hour >= 0 && hour <= 9)
	{
		hour = "0" + hour;
	}
	if (minutes >= 0 && minutes <= 9)
	{
		minutes = "0" + minutes;
	}
	if (seconds >= 0 && seconds <= 9)
	{
		seconds = "0" + seconds;
	}

	var currentdate = hour + sign2 + minutes + sign2 + seconds;
	
	return currentdate;
}

function ALL(){
	var b = new Boolean();
	b = true;
	for(var i=0;i< arguments.length; ++i)
	{
		if(!arguments[i])
		{
			b = false;
			break;
		}	
	}
	return b;
}

function ANY(){
	var b = new Boolean();
	b = false;
	for(var i=0;i< arguments.length; ++i)
	{
		if(arguments[i])
		{
			b = true;
			break;
		}
	}
	return b;
}

//---------------------------------------------------  
// 日期格式化  
// 格式 YYYY/yyyy/YY/yy 表示年份  
// MM/M 月份  
// dd/DD/d/D 日期  
// hh/HH/h/H 时间  
// mm/m 分钟  
// ss/SS/s/S 秒  
//---------------------------------------------------  
function STRTODATE(date,formatStr){
	
	var d = date;
	/*新模式预览时时区不正确
	var timeoffset = (new Date().getTimezoneOffset()) * 60 * 1000;
	var len = d.getTime();
	d = new Date(len - timeoffset);
	*/
//	d = date2.getFullYear() + "/" + (date2.getMonth() + 1) + "/" + date2.getDate();
//	var d = new Date(date);
	var str = formatStr;   
//    var Week = ["Sun","Mon","Tue","Wed","Thu","Fri","Sat"];
  
    str=str.replace(/yyyy|YYYY/,d.getFullYear());   
    str=str.replace(/yy|YY/,(d.getYear() % 100)>9?(d.getYear() % 100).toString():'0' + (d.getYear() % 100));   
  
    str=str.replace(/MM/,d.getMonth()+1>9?(d.getMonth()+1).toString():'0' + (d.getMonth()+1));   
    str=str.replace(/M/g,d.getMonth()+1);
  
//    str=str.replace(/w|W/g,Week[d.getDay()]);   
  
    str=str.replace(/dd|DD/,d.getDate()>9?d.getDate().toString():'0' + d.getDate());   
    str=str.replace(/d|D/g,d.getDate());   
  
    str=str.replace(/hh|HH/,d.getHours()>9?d.getHours().toString():'0' + d.getHours());   
    str=str.replace(/h|H/g,d.getHours());   
    str=str.replace(/mm/,d.getMinutes()>9?d.getMinutes().toString():'0' + d.getMinutes());   
    str=str.replace(/m/g,d.getMinutes());   
  
    str=str.replace(/ss|SS/,d.getSeconds()>9?d.getSeconds().toString():'0' + d.getSeconds());   
    str=str.replace(/s|S/g,d.getSeconds());   
  
    return str;   
}

function TOPRECISION(value,num){
	var n = new Number(value);
	return n.toPrecision(num);
}

function TOFIXED(value,num){
	var n = new Number(value);
	return n.toFixed(num);
}

function TOEXPONENTIAL(value,num){
	var n = new Number(value);
	return n.toExponential(num);
}

function CHARAT(value,index){
	var string = new String(value);
	return string.charAt(index);
}

function PUSH(value,newelement1){
	var arr = value;
	arr.push(newelement1);
	return arr;
}

function SUM(...value){
	var n = 0;
	var newArr = [];
	arrConversion(newArr,value);
	for(var i=0;i<newArr.length;++i){
		n += newArr[i];
	}	
	return n;
}

function REVERSE(arr){
	/*
	var arr = new Array;
	for(var i=0;i< arguments.length; ++i)
	{
		var array = arguments[i];
		arr.push(array);
	}*/
	return arr.reverse();
}

function LENGTH(object){
	return object.length;
}

function SORTED(arr){
	/*
	var arr = new Array;
	for(var i=0;i< arguments.length; ++i)
	{
		var array = arguments[i];
		arr.push(array);
	}*/
	return arr.sort();
}

function CONTAINS(value,substring)
{
	var string = new String(value);
	var index = string.indexOf(substring);
	var b = (index != -1) ? true : false; 
	return b;
}

function NUMATARRAYCOUNT(arr, num) {
    arr.sort();
	var lastIndex = arr.lastIndexOf(num);
	var firstIndex = arr.indexOf(num);
    return (lastIndex == -1) ? 0 : (lastIndex - firstIndex + 1);
}

function FILTERARRAY(arr,value){
	var len = arr.length;
	var filterArr = new Array;
	var remainValue = 0;
	var bFlag = false;
	for( var i=0;i<len;++i){
		var v = arr[i];
		if(v > value){
			filterArr.push(v);
		}else{
			bFlag = true;
			remainValue += v;
		}
	}
	if(bFlag){
		filterArr.push(remainValue);
	}
	return filterArr;
}

function ARRAYMAP(arrValue,arrS,arrD){
	var arrValueLen = arrValue.length;
	var arrDLen  = arrS.length;
    var arrReturn = new Array;
	var pos = 0;
	for(var i=0;i<arrDLen ;++i){
		var v = arrD[i];
		var index = arrS.indexOf(v,pos);
		if(index >= 0 && index < arrValueLen){
			arrReturn.push(arrValue[index]);
			pos = index + 1;
		}		
	}
	return arrReturn;
}

function FILTERVALUE(arr,value){
	var len = arr.length;
	var filterArr = new Array;

	for( var i=0;i<len;++i){
		var v = arr[i];
		if(v < value){
			filterArr.push(v);
		}
	}
	
	return filterArr;
}

function FILTERARRAYSTR(arr,valueStr){
	var len = arr.length;
	var filterArr = new Array;
	for( var i=0;i<len;++i){
		var v = arr[i];
		if(v != valueStr){
			filterArr.push(v);
		}
	}

	return filterArr;
}

function STRARRAY(arr){
	return arr.split(',');
}

function DAYSDIFF(DateOne,DateTwo)
{
	var d1=new Date(DateOne);
	var d2=new Date(DateTwo);
	
	var cha=(d2.getTime()-d1.getTime())/86400000; 
	return cha;
}

function ADDDATE(date,days,formatStr){ 
    var d=new Date(date); 
    d.setDate(d.getDate()+days); 
    return STRTODATE(d,formatStr);
}

function SUBDATE(date,days,formatStr){ 
    var d=new Date(date); 
    d.setDate(d.getDate()-days); 
    return STRTODATE(d,formatStr);
}

function CURMONTHDAYS(){
	var now=new Date();
	var d = new Date(now.getFullYear(),now.getMonth()+1,0);
	return d.getDate();
}

function GETMONTHDAYS(year, month){
    var d = new Date(year, month, 0);
    return d.getDate();
}

function CURMONTHFIRSTDAY(formatStr){
	var now=new Date();
	var d = new Date(now.getFullYear(),now.getMonth(),1);
	return STRTODATE(d,formatStr);
}

function CURMONTHLASTDAY(formatStr){
	var now=new Date();
	var d = new Date(now.getFullYear(),now.getMonth()+1,0);
	return STRTODATE(d,formatStr);
}

function GETMONTHFIRSTDAY(year, month,formatStr){
	var d = new Date(year, month-1, 1);
	return STRTODATE(d,formatStr);
}

function NEWGUID(){
    var guid = "";
    for (var i = 1; i <= 32; i++){
      var n = Math.floor(Math.random()*16.0).toString(16);
      guid +=   n;
      if((i==8)||(i==12)||(i==16)||(i==20))
        guid += "-";
    }
    return guid;    
}

function PARSEFLOAT(string){
	return parseFloat(string);
}

function PARSEINT(string, radix){
	return parseInt(string, radix);
}

function EVAL(string){
	return eval(string);
}


//-----------以下EXCEL函数--------------//

function ABS(number) {
	return formulajs.ABS(number);
};

function ACOS(number) {
	return formulajs.ACOS(number);
};

function ACOSH(number) {
	return formulajs.ACOSH(number);
};

function ACOT(number) {
	return formulajs.ACOT(number);
};

function ACOTH(number) {
	return formulajs.ACOTH(number);
};

function AND() {
	return formulajs.AND(...arguments);
};

function ASIN(number) {
	return formulajs.ASIN(number);
};

function ASINH(number) {
	return formulajs.ASINH(number);
};

function ATAN(number) {
	return formulajs.ATAN(number);
};

function ATAN2(number_x, number_y) {
	return formulajs.ATAN2(number_x, number_y);
};
	
function ATANH(number) {
	return formulajs.ATANH(number);
};

function AVERAGE() {
	return formulajs.AVERAGE(...arguments);
};

function AVERAGEIF(range, criteria, average_range) {
	return formulajs.AVERAGEIF(range, criteria, average_range);
}

function BASE(number, radix, min_length) {
	return formulajs.BASE(number, radix, min_length);
};

function BIN2DEC(number) {
	return formulajs.BIN2DEC(number);
}

function BIN2HEX(number, places) {
	return formulajs.BIN2HEX(number, places);
}

function BIN2OCT(number, places) {
	return formulajs.BIN2OCT(number, places);
}

function BITAND(number1, number2) {
	return formulajs.BITAND(number1, number2);
}

function BITLSHIFT(number, shift) {
	return formulajs.BITLSHIFT(number, shift);
}

function BITOR(number1, number2) {
	return formulajs.BITOR(number1, number2);
}

function BITRSHIFT(number, shift) {
	return formulajs.BITRSHIFT(number, shift);
}

function BITXOR(number1, number2) {
	return formulajs.BITXOR(number1, number2);
}

function CEILING(number, significance) {
	var mode = significance > 0 ? 0 : -1;
	return formulajs.CEILING(number, significance,mode);
};

function CHAR(number) {
	return formulajs.CHAR(number);
};

function CHOOSE() {
	return formulajs.CHOOSE(...arguments);
}

function CODE(number) {
	return formulajs.CODE(number);
};

function COLUMN(reference) {
	var n = 0;
	var s = reference.replace(/[^a-zA-Z]/g,'');
	var index = reference.indexOf(s);
	if(index !=0){
		return 0;
	}
	var j = 0;
	for (var i = s.length - 1, j = 1; i >= 0; i--, j *= 26) {
		var c = s[i].toUpperCase();
		if (c < 'A' || c > 'Z') {
			return 0;
		}
		n += (c.charCodeAt(0) - 64) * j;
	}
	return n;
}

function COMBIN(number, number_chosen){
	return formulajs.COMBIN(number, number_chosen);
}

function COMBINA(number, number_chosen){
	return formulajs.COMBINA(number, number_chosen);
}

function COMPLEX(real, imaginary, suffix){
	return formulajs.COMPLEX(real, imaginary, suffix);
}

function CONCATENATE() {
	return formulajs.CONCATENATE(...arguments);
};

function CONVERT(number, from_unit, to_unit) {
	return formulajs.CONVERT(number, from_unit, to_unit);
}

function COS(number){
	return formulajs.COS(number);
}

function COSH(number){
	return formulajs.COSH(number);
}

function COT(number){
	return formulajs.COT(number);
}

function COTH(number){
	return formulajs.COTH(number);
}

function COUNT() {
	return formulajs.COUNT(...arguments);
}

function CSC(number){
	return formulajs.CSC(number);
}

function CSCH(number){
	return formulajs.CSCH(number);
}

function DEC2BIN(number, places) {
	return formulajs.DEC2BIN(number, places);
}

function DEC2HEX(number, places) {
	return formulajs.DEC2HEX(number, places);
}

function DEC2OCT(number, places) {
	return formulajs.DEC2OCT(number, places);
}

function DELTA(number1, number2) {
	return formulajs.DELTA(number1, number2);
}

function DOLLAR(number, decimals){
	return formulajs.DOLLAR(number, decimals);
}

function DATE(year, month, day) {
	return formulajs.DATE(year, month, day);
};

function DATEVALUE(date_text) {
	return formulajs.DATEVALUE(date_text);
}

function DAY(serial_number) {
	return formulajs.DAY(serial_number);
};

function DAYS(end_date, start_date) {
	return formulajs.DAYS(end_date, start_date);
}

function DAYS360(start_date, end_date, method=false) {
	return formulajs.DAYS360(start_date, end_date, method);
}

function DECIMAL(number, radix){
	return formulajs.DECIMAL(number, radix);
}

function DEGREES(number){
	return formulajs.DEGREES(number);
}

function EDATE(start_date, months){
	return formulajs.EDATE(start_date, months);
}

function EOMONTH(start_date, months) {
	return formulajs.EOMONTH(start_date, months);
}

function EXACT(text1, text2){
	return formulajs.EXACT(text1, text2);
}

function EVEN(number){
	return formulajs.EVEN(number);
}

function FACT(number){
	return formulajs.FACT(number);
}

function FACTDOUBLE(number){
	return formulajs.FACTDOUBLE(number);
}

function FALSE(){
	return formulajs.FALSE()
}

function FIND(find_text, within_text, position){
	return formulajs.FIND(find_text, within_text, position);
}

function FIXED(number, decimals, no_commas=true){
	return formulajs.FIXED(number, decimals, no_commas);
}

function FLOOR(number, significance){
	var mode = 0;
	if(significance < 0){
		mode = -1;
	}
	return formulajs.FLOOR.MATH(number, significance,mode);
}

function GCD(){
	return formulajs.GCD(...arguments);
}

function GESTEP(number, step) {
	return formulajs.GESTEP(number, step);
}

function HEX2BIN(number, places) {
	return formulajs.HEX2BIN(number, places);
}

function HEX2DEC(number, places) {
	return formulajs.HEX2DEC(number, places);
}

function HEX2OCT(number, places) {
	return formulajs.HEX2OCT(number, places);
}

function HOUR(serial_number){
	return formulajs.HOUR(serial_number);
}

function IF(cons,value1,value2){
	return formulajs.IF(cons,value1,value2);
}

function IMABS(inumber){
	return formulajs.IMABS(inumber);
}

function IMAGINARY(inumber) {
	return formulajs.IMAGINARY(inumber);
}

function IMCONJUGATE(inumber) {
	return formulajs.IMCONJUGATE(inumber);
}

function IMDIV(inumber1, inumber2) {
	return formulajs.IMDIV(inumber1, inumber2);
}

function IMREAL(inumber) {
	return formulajs.IMREAL(inumber);
}

function IMSUB(inumber1, inumber2) {
	return formulajs.IMSUB(inumber1, inumber2);
}

function IMSUM() {
	return formulajs.IMSUM(...arguments);
}

function INDEX(array, colum_count, row_num, column_num) {
	if(column_num > colum_count){
		return undefined;
	}
	var pos = colum_count*(row_num-1)+column_num;
	return array[pos-1];
}

function INT(number) {
	return formulajs.INT(number);
};

function ISOWEEKNUM(date) {
	return formulajs.ISOWEEKNUM(date);
}

function LCM(){
	return formulajs.LCM(...arguments);
}

function LEFT(text, number) {
	return formulajs.LEFT(text, number);
};

function LEN(text) {
	return formulajs.LEN(text);
};

function LN(number) {
	return formulajs.LN(number);
};

function LOG(number, base =10){
	return formulajs.LOG(number, base);
}

function LOG10(number){
	return formulajs.LOG10(number);
}

function LOWER(text) {
	return formulajs.LOWER(text);
}

function MATCH(lookupValue, lookupArray, matchType) {
	return formulajs.MATCH(lookupValue, lookupArray, matchType);
};

function MAX() {
	return formulajs.MAX(...arguments);
};

function MDETERM(matrix) {
	return formulajs.MDETERM(matrix);
}

function MID(text, start, number) {
	return formulajs.MID(text, start, number);
};

function MIN() {
	return formulajs.MIN(...arguments);
};

function MINUTE(serial_number) {
	return formulajs.MINUTE(serial_number);
}

function MOD(dividend, divisor) {
	return formulajs.MOD(dividend, divisor);
};
	
function MODE() {
	return formulajs.MODE(...arguments);
};

function MONTH(serial_number) {
	return formulajs.MONTH(serial_number);
};

function MROUND(number, multiple) {
	return formulajs.MROUND(number, multiple);
}

function MULTINOMIAL(){
	return formulajs.MULTINOMIAL(...arguments);
}

function MUNIT(dimension){
	return formulajs.MUNIT(dimension);
}

function NETWORKDAYS(start_date, end_date, holidays){
	return formulajs.NETWORKDAYS(start_date, end_date, holidays);
}

function NOT(logical){
	return formulajs.NOT(logical);
}

function NOW() {
	return formulajs.NOW();
};

function OCT2BIN(number, places){
	return formulajs.OCT2BIN(number, places);
}

function OCT2DEC(number, places){
	return formulajs.OCT2DEC(number, places);
}

function OCT2HEX(number, places){
	return formulajs.OCT2HEX(number, places);
}

function ODD(number) {
	return formulajs.ODD(number);
}

function OR() {
	return formulajs.OR(...arguments);
};

function PI() {
	  return formulajs.PI();
};

function POWER(number, power){
	return formulajs.POWER(number, power);
}

function PRODUCT() {
	return formulajs.PRODUCT(...arguments);
}

function PROPER(text) {
	return formulajs.PROPER(text);
}

function QUOTIENT(numerator, denominator) {
	return formulajs.QUOTIENT(numerator, denominator);
}

function RADIANS(number) {
	return formulajs.RADIANS(number);	
}

function RAND(){
	return Math.random();
}

function RANDBETWEEN(bottom, top){
	return formulajs.RANDBETWEEN(bottom, top);	
}

function RANK(number, range, order) {
	return formulajs.RANK(number, range, order);
};

function REPLACE(text, position, length, new_text) {
	return formulajs.REPLACE(text, position, length, new_text);
}

function REPT(text, number) {
	return formulajs.REPT(text, number);
}

function RIGHT(text, number) {
	return formulajs.RIGHT(text, number);
};

function ROUND(number, digits) {
	return formulajs.ROUND(number, digits);
};

function ROUNDDOWN(number, digits){
	return formulajs.ROUNDDOWN(number, digits);
}

function ROUNDUP(number, digits){
	return formulajs.ROUNDUP(number, digits);
}

function SEARCH(find_text, within_text, position) {
	return formulajs.SEARCH(find_text, within_text, position);
}

function SEC(number) {
	return formulajs.SEC(number);
}

function SECH(number) {
	return formulajs.SECH(number);
}

function SECOND(serial_number) {
	return formulajs.SECOND(serial_number);
}

function SIGN(number) {
	return formulajs.SIGN(number);
}

function SIN(number) {
	return formulajs.SIN(number);
}

function SINH(number) {
	return formulajs.SINH(number);
}

function SQRT(number) {
	return formulajs.SQRT(number);
}

function SQRTPI(number) {
	return formulajs.SQRTPI(number);
}

function SUBSTITUTE(text, old_text, new_text, occurrence) {
	return formulajs.SUBSTITUTE(text, old_text, new_text, occurrence);
}

function SUM() {
	return formulajs.SUM(...arguments);
}
	
function SUMIF(range, criteria) {
	return formulajs.SUMIF(range, criteria);
}

function SUMPRODUCT() {
	return formulajs.SUMPRODUCT(...arguments);
}

function SUMSQ() {
	return formulajs.SUMSQ(...arguments);
}

function T(value) {
	return (typeof value === "string") ? value : '';
};

function TAN(number) {
	return formulajs.TAN(number);
}

function TANH(number) {
	return formulajs.TANH(number);
}

function TIME(hour, minute, second){
	return formulajs.TIME(hour, minute, second);
}

function TIMEVALUE(time_text) {
	return formulajs.TIMEVALUE(time_text);
}

function TODAY() {
	return formulajs.TODAY();
};

function TRIM(text) {
	return formulajs.TRIM(text);
};

function TRUE(){
	return formulajs.TRUE();
}

function TRUNC(number, digits){
	return formulajs.TRUNC(number, digits);
}

function TYPE(value) {
	return formulajs.TYPE(value);
}

function UPPER(text) {
	return formulajs.UPPER(text);
}

function VALUE(text) {
	return formulajs.VALUE(text);
}

function WEEKDAY(serial_number, return_type){
	return formulajs.WEEKDAY(serial_number, return_type);
}

function WEEKNUM(serial_number, return_type){
	return formulajs.WEEKNUM(serial_number, return_type);
}

function WORKDAY(start_date, days, holidays){
	return formulajs.WORKDAY(start_date, days, holidays);
}

function XOR(){
	return formulajs.XOR(...arguments);
}

function YEAR(serial_number) {
	return formulajs.YEAR(serial_number);
}

function YEARFRAC(start_date, end_date, basis) {
	return formulajs.YEARFRAC(start_date, end_date, basis);
}

function DATEDIFF(date,day,strformat){
	if(strformat == undefined){
		strformat ='YYYY-MM-DD HH:mm:ss';
	}
	var d = dayjs(date);
	d = d.add(day, 'day');
	return d.format(strformat);
}

function ASSERTDIVZERO(a) {
	return isFinite(a)?a:NaN;
}

function COLWIDTH(colIndex) {
	return efreportJS.getColWidth(colIndex);
}

function ROWHEIGHT(rowIndex) {
	return efreportJS.getRowHeight(rowIndex);
}

function TODAYWEEKDAY(){
	var weekday = ["星期日","星期一","星期二","星期三","星期四","星期五","星期六"]
	var d = dayjs().day();
	return weekday[d];
}

function MONTH2CHINESE(){
	var chinese = ["一","二","三","四","五","六","七","八","九","十","十一","十二"]
	var m = dayjs().month();
	return chinese[m];
}

//特殊自定义函数
function BBDATE(paramname , year , month , day , hour , minute){
	var date = dayjs(paramname);
	if(year == "T"){
	}else if(year == "L"){
		date = date.add(-1, 'year');
	}
	
	if(month == "L"){
		date = date.add(-1, 'month');
	} 
	else if (month == "T") {
	}
	else if (month == "N"){
		date = date.add(1, 'month');
	}else {
		var mon = parseInt(month) -1;
		date = date.set('month', mon);
	}
	
	if(day == "L"){
		date = date.add(-1, 'day');	
	} else if (day == "T"){		
	}else if (day == "N"){
		date = date.add(1, 'day');
	}else {
		var d = parseInt(day);
		date = date.set('date', d);
	}
	
	var h = parseInt(hour);
	date = date.set('hour',h);
	
	var m = parseInt(minute);
	date = date.set('minute',m);
	
	date = date.set('second',0);
	
	return date.format('YYYY-MM-DD HH:mm:ss');
}

function DATACONVERT(arr,grouplenght){		
	var arrlength = arr.length;
	var newArr = [];
	for(var j=0;j<grouplenght;j++){
		for(var i=0;i<arrlength/grouplenght;i++){
			newArr.push(arr[i*grouplenght+j]);
			}
	}
	return newArr;
}

function FIXEDARRAY(value,num){
	var newArr = [];
	for(var i=0;i<value.length;++i){
		newArr.push(FIXED(value[i],num));
	}
	return newArr;
}

//非字段数据整合成符合图表的数组
function FIXEDDATACONVERT(arr,grouplenght,num){		
	var arrlength = arr.length;
	var newArr = [];
	for(var j=0;j<grouplenght;j++){
		for(var i=0;i<arrlength/grouplenght;i++){
			newArr.push(FIXED(arr[i*grouplenght+j],num));
			}
	}
	return newArr;
}

//第一个参数是字段数组，第二个参数是非字段数组，最后合并成符合图表的数组
function CONVERTDATABYARRAY(arr1,arr2,grouplenght,num){		
	var arrlength1 = arr1.length;
	var arrnew2 = FIXEDDATACONVERT(arr2,grouplenght,num);
	var arrlength2 = arr2.length;
	var group1 = arrlength1/grouplenght;
	var group2 = arrlength2/grouplenght;
	var pos1 = 0;
	var pos2 = 0;
	var newArr = [];
	for(var j=0;j<grouplenght;j++){	
		for(var i=0;i<group1;++i){
			newArr.push(FIXED(arr1[pos1++],num));
		}
		for(var i=0;i<group2;++i){
			newArr.push(arrnew2[pos2++]);
		}
	}
	return newArr;
}

function LASTDATA(arr,num){
	var lst = 0;
	var length = arr.length;
	lst = arr.pop();
	while(lst == num){
		lst = arr.pop();
	}
	return lst == undefined ? 0 : lst;
}

function ARRAYINIT(leng,num){
	var array = [];
	for(var i=0;i<leng;++i){
		array.push(num)
	}
	return array;
}

var globalCount = 0;

function AUTOINC(count){
	if(count != undefined){
		globalCount = count;
		return;
	}
	globalCount++;
	return globalCount;
}

function ISHEADSTRING(value,substring){
	var string = new String(value);
	var index = string.indexOf(substring);
	var b = (index == 0) ? true : false; 
	return b;
}

function ISLASTSTRING(value,substring){
	var string = new String(value);
	var d = string.length - substring.length;
	return (d >= 0 && string.lastIndexOf(substring) == d)
}

function ISCONTAINSTRARRAY(srcStr, str){
	var v;
	if(typeof srcStr === 'number' && !isNaN(srcStr)){
		v = srcStr.toString();
	}else{
		v = srcStr;
	}
	var ar = str.split(',');
	var index = ar.indexOf(v);
	return (index > -1);
}

function INTTOSTR_27_I(x)
{
    if (x < 0)
        return "";

    var a = x / 26;
    var b = x % 26;
    return INTTOSTR_27_I(a - 1) + String.fromCharCode(b + 'A'.charCodeAt());
}

//数字转换成A-Z
function ITOS_27(x)
{
    return INTTOSTR_27_I(x - 1);
}

//A-Z转换成数字
function STOI_27(index)
{
    if (1 == index.length) {
        return index.charAt(0).charCodeAt() - 64;
    }

    var a = 0;
    var count = index.length;
    var i = 0;
    for (i = 0; i < count; i++) {
        a += ((index.charAt(i).charCodeAt()) - 64) * Math.pow(26, count - i - 1);
    }

    return a;
}

function MATCHVALUE(lookupValue,lookupArray,mapArray, defaultValue){
	if(lookupArray.length != mapArray.length || lookupArray.leng == 0 || mapArray.length == 0){
		return defaultValue; 
	}
	
	var index = lookupArray.indexOf(lookupValue);
	if(index == -1){
		return defaultValue; 
	}
	
	return mapArray[index];
}

function REMOVEDATASOURCE(dsName) {
	return efreportJS.removeDataSource(dsName);
}

//�ݴ�ű�
function ResumeError()
 {
        return true;
    }
window.onerror = ResumeError;

//����Ҽ����Խ�ֹ��
/*
if (window.Event) 
document.captureEvents(Event.MOUSEUP); 
function nocontextmenu() 
{
event.cancelBubble = true
event.returnValue = false;
return false;
}
function norightclick(e) 
{
if (window.Event) 
{
if (e.which == 2 || e.which == 3)
return false;
}
else
if (event.button == 2 || event.button == 3)
{
event.cancelBubble = true
event.returnValue = false;
return false;
}
}
document.oncontextmenu = nocontextmenu; // for IE5+
document.onmousedown = norightclick; // for all others
*/

 //����Ƿ������ַ�
  function is_zw(str)
{
	exp=/[0-9a-zA-Z_.,#@!$%^&*()-+=|\?/<>]/g;
	if(str.search(exp) != -1)
	{
		return false;
	}
	return true;
}
//��֤�Ƿ��������
function CheckBadChar(Obj,AlertStr)
{
	exp=/[,��]/g;
	if(Obj.value.search(exp) != -1)
	{   alert(AlertStr+"���ܰ�������");
	    Obj.value="";
		Obj.focus();
		return false;
	}
	return true;
}
// ����Ƿ���Ч����չ��
function IsExt(FileName, AllowExt){
		var sTemp;
		var s=AllowExt.toUpperCase().split("|");
		for (var i=0;i<s.length ;i++ ){
			sTemp=FileName.substr(FileName.length-s[i].length-1);
			sTemp=sTemp.toUpperCase();
			s[i]="."+s[i];
			if (s[i]==sTemp){
				return true;
				break;
			}
		}
		return false;
}
//����Ƿ����ַ���һ
function is_number(str)
{
	exp=/[^0-9()-]/g;
	if(str.search(exp) != -1)
	{
		return false;
	}
	return true;
}
//������ַ�����
function CheckNumber(Obj,DescriptionStr)
{
	if (Obj.value!='' && (isNaN(Obj.value) || Obj.value<0))
	{
		alert(DescriptionStr+"Ӧ����Ч���֣�");
		Obj.value="";
		Obj.focus();
		return false;
	}
	return true;
}
//�������ʼ���Ч��
function is_email(str)
{ if((str.indexOf("@")==-1)||(str.indexOf(".")==-1)){
	
	return false;
	}
	return true;
}
function CheckAll(form)
{
				  for (var i=0;i<form.elements.length;i++)
				  {
					var e = form.elements[i];
					if (e.Name != 'chkAll'&&e.disabled==false)
					   e.checked = form.chkAll.checked;
					}
 } 
function OpenWindow(Url,Width,Height,WindowObj)
{
	var ReturnStr=showModalDialog(Url,WindowObj,'dialogWidth:'+Width+'pt;dialogHeight:'+Height+'pt;status:no;help:no;scroll:no;status:0;help:0;scroll:0;');
	return ReturnStr;
}
function WinPop(url, width, height)
{
  window.showModelessDialog(url,"",'dialogWidth=' + width + 'px; dialogHeight=' + height + 'px; resizable=no; help=no; scroll=no; status=no;resizable=0; help=0; scroll=0; status=0;'); 
}
function OpenThenSetValue(Url,Width,Height,WindowObj,SetObj)
{
	var ReturnStr=showModalDialog(Url,WindowObj,'dialogWidth:'+Width+'pt;dialogHeight:'+Height+'pt;status:no;help:no;scroll:no;status:0;help:0;scroll:0;');
	if (ReturnStr!='') SetObj.value=ReturnStr;
	return ReturnStr;
}
function OpeneditorWindow(Url,WindowName,Width,Height)
{
	window.open(Url,WindowName,'toolbar=0,location=0,maximize=1,directories=0,status=1,menubar=0,scrollbars=0,resizable=1,top=50,left=50,width='+Width+',height='+Height);
}

function CheckEnglishStr(Obj,DescriptionStr)
{
	var TempStr=Obj.value,i=0,ErrorStr='',CharAscii;
	if (TempStr!='')
	{
		for (i=0;i<TempStr.length;i++)
		{
			CharAscii=TempStr.charCodeAt(i);
			if (CharAscii>=255||CharAscii<=31)
			{
				ErrorStr=ErrorStr+TempStr.charAt(i);
			}
			else
			{
				if (!CheckErrorStr(CharAscii))
				{
					ErrorStr=ErrorStr+TempStr.charAt(i);
				}
			}
		}
		if (ErrorStr!='')
		{
			alert("������Ϣ:\n\n"+DescriptionStr+'���ַǷ��ַ�:'+ErrorStr);
			Obj.focus();
			return false;
		}
		if (!(((TempStr.charCodeAt(0)>=48)&&(TempStr.charCodeAt(0)<=57))||((TempStr.charCodeAt(0)>=65)&&(TempStr.charCodeAt(0)<=90))||((TempStr.charCodeAt(0)>=97)&&(TempStr.charCodeAt(0)<=122))))
		{
			alert(DescriptionStr+'���ַ�ֻ�ܹ�Ϊ���ֻ�����ĸ');
			Obj.focus();
			return false;
		}
	}
	return true;
}
function CheckErrorStr(CharAsciiCode)
{
	var TempArray=new Array(34,47,92,42,58,60,62,63,124);
	for (var i=0;i<TempArray.length;i++)
	{
		if (CharAsciiCode==TempArray[i]) return false;
	}
	return true;
}
//Obj�����Ķ���,OpStr--BottomFrame��ʾ��ǰ��������ʾ��Ϣ,ButtonSymbol��ť״̬,MainUrl--MainFrame������
function SelectObjItem(Obj,OpStr,ButtonSymbol,MainUrl)
{   if (OpStr!='')
    {window.parent.parent.frames['BottomFrame'].location.href='../KS.Split.asp?OpStr='+OpStr+'&ButtonSymbol='+ButtonSymbol;}
	if(MainUrl!='')
	{window.parent.parent.frames['MainFrame'].location.href=MainUrl;
	}
	if (Obj!='')
	 {
	   for (var i=0;i<document.all.length;i++)
	   {
		if (document.all(i).className=='FolderSelectItem') document.all(i).className='FolderItem';
	    }
	   Obj.className='FolderSelectItem';
	}
}
//Obj�����Ķ���,OpStr--BottomFrame��ʾ��ǰ��������ʾ��Ϣ,ButtonSymbol��ť״̬,MainUrl--MainFrame������
function SelectObjItem1(Obj,OpStr,ButtonSymbol,MainUrl,ChannelID)
{   if (OpStr!='')
    {window.parent.parent.frames['BottomFrame'].location.href='KS.Split.asp?ChannelID='+ChannelID+'&OpStr='+OpStr+'&ButtonSymbol='+ButtonSymbol;}
	if(MainUrl!='')
	{window.parent.parent.frames['MainFrame'].location.href=MainUrl;
	}
	//if (Obj!='')
	// {
	//   for (var i=0;i<document.all.length;i++)
	//   {
	//	if (document.all(i).className=='FolderSelectItem') document.all(i).className='FolderItem';
	//    }
	//   Obj.className='FolderSelectItem';
	//}
}
function FolderClick(Obj,el)
{   	var i=0;
  for (var i=0;i<document.all.length;i++)
	   {
		if (document.all(i).className=='FolderSelected') document.all(i).className='';
	    }
	         Obj.className='FolderSelected';
	  
              for (i=0;i<DocElementArr.length;i++)
			{
				if (el==DocElementArr[i].Obj)
				{
					if (DocElementArr[i].Selected==false)
					{
						DocElementArr[i].Obj.className='FolderSelectItem';
						DocElementArr[i].Selected=true;
					}
					else
					{
						DocElementArr[i].Obj.className='FolderItem';
						DocElementArr[i].Selected=false;
					}
				}
			}
}
function InsertKeyWords(obj,KeyWords)
{
	if (KeyWords!='')
	{
		if (obj.value.search(KeyWords)==-1)
		{
			if (obj.value=='') obj.value=KeyWords;
			else obj.value=obj.value+'|'+KeyWords;
			
		}
	}
	if (KeyWords == 'Clean')
	{
		obj.value = '';
	}
	return;
}
function Getcolor(img_val,Url,input_val){
	var arr = showModalDialog(Url, "", "dialogWidth:18.5em; dialogHeight:17.5em; status:0; help:0");
	if (arr != null){
		document.getElementById(input_val).value = arr;
		img_val.style.backgroundColor = arr;
		}
}
//���Ͳ���������Frames����
function SendFrameInfo(MainUrl,LeftUrl,ControlUrl)
{
	location.href=MainUrl;
    parent.frames['LeftFrame'].LeftInfoFrame.location.href=LeftUrl;
	 parent.frames['BottomFrame'].location.href=ControlUrl;
}
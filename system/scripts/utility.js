var Browser = new Object();
Browser.ua = getUserAgent();
function getUserAgent()
{
    var ua = navigator.userAgent.toLowerCase();

    if (ua.indexOf("opera") >= 0)
    {
        return "opera";
    }

    if (ua.indexOf("firefox") >= 0)
    {
        return "ff";
    }

    if (ua.indexOf("gecko") >= 0)
    {
        return "moz";
    }

    if (ua.indexOf("msie"))
    {
        ieVer = parseFloat(ua.substr(ua.indexOf("msie") + 5));

        if (ieVer >= 6)
        {
            return "ie6";
        }

        if (ieVer >= 5.5)
        {
            return "ie55";
        }

        if (ieVer >= 5)
        {
            return "ie5";
        }
    }

    return "other";
}

function setCookie(name, value, expires, path, domain)
{
    var str = name + "=" + escape(value);

    if (expires)
    {
        if (expires == 'never')
        {
            expires = 100 * 365 * 24 * 60;
        }

        var exp = new Date();
        exp.setTime(exp.getTime() + expires * 60 * 1000);
        str += "; expires=" + exp.toGMTString();
    }

    if (path)
    {
        str += "; path=" + path;
    }

    if (domain)
    {
        str += "; domain=" + domain;
    }

    document.cookie = str;
}

function getCookie(name)
{
    var tmp, reg = new RegExp("(^| )" + name + "=([^;]*)(;|$)", "gi");
    tmp = reg.exec(unescape(document.cookie));

    if (tmp)
    {
        return (tmp[2]);
    }

    return null;
}

function deleteusercookie(name) {
	path='';
	domain='localhost';
    if (getCookie(name)) {
        document.cookie = name + "=" +
            ((path) ? "; path=" + path : "") +
            ((domain) ? "; domain=" + domain : "") +
            "; expires=Thu, 01-Jan-70 00:00:01 GMT";
    }
}

function getFullUrl(url)
{
    return (url.indexOf('http://') == 0 || url.indexOf('https://') == 0) ? url : (url.indexOf('/') == 0)
        ? location.protocol + '//' + location.host + url : (url.indexOf('www') == 0)
        ? 'http://' + url : location.href.substr(0, location.href.lastIndexOf('/') + 1) + url;
}

function getTureLength(str)
{
    if (str == null)
    {
        return null;
    }

    var PatSWord = /^[\x00-\xff]+$/;
    var PatDWord = /[^\x00-\xff]+/g;
    var ln = 0;

    for (var i = 0; i < str.length; i++)
    {
        var char = str.charAt(i);

        if (PatSWord.test(char))
        {
            ln += 1;
        }
        else
        {
            ln += 2;
        }
    }

    return ln;
}


function getStrLength(str)
{
	var i = 0;
	var j = 0;

	for (i = 0; i < str.length; i++)
	{
		if (str.charCodeAt(i) > 127 || str.charCodeAt(i) == 94)
		{
			j = j + 2;
		}
		else
		{
			j = j + 1;
		}
	}

	return j;
}


function isEmail(email)
{
    var rStr = new RegExp("[^a-z,0-9,_,--,@,\.]", "ig");

    if ((!email.match(rStr)) && email.length > 5 && email.indexOf('@') > 0 && email.indexOf('.') > 0)
        return true;
    else
        return false;
}
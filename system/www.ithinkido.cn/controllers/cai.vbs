Class Cai
	Function index()

		Set Conn = Server.CreateObject("ADODB.Connection") 
		DBPath = Server.MapPath("laoxu.mdb") 
		Conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="&DBPath
		'���������ݿ�����

		ye="3"
		'�ӵ�һҳ�����ڼ�ҳ��������3


		spic="1"
		'�������Ƿ񱣴�ͼƬ"0"Ϊ����"1"Ϊ����������汾���ṩ

		flashsw = "1"
		'�������Ƿ񱣴�flash"0"Ϊ����"1"Ϊ����������汾���ṩ

		const saveflash="../upload/flash/" 'flash����·������汾���ṩ

		const savepath="../upload/flash/" 'ͼƬ����·������汾���ṩ


		dim rs,t,r,rs1,sql,url,StrId,overId,body_str,body_str_str,next_str,flash_url,flash_name,flash_zuozhe
		if request("url") = "" then
				response.write "<script language='javascript'>" & VbCRlf
				response.write "alert('��������');" & VbCrlf
				response.write "history.go(-1);" & vbCrlf
				response.write "</script>" & VbCRLF
				response.end
		end if 
		r=12
		for i = ye To 1 Step -1

			  
			  if i=1 then
			  url="http://flash.86516.com/type.asp?cd=%D2%F4%C0%D6&offset=0"
			  else
			  t=(i-1)*r
			  url="http://flash.86516.com/type.asp?cd=%D2%F4%C0%D6&offset="&t
			  end if
			  
			  url = replace(url,"{id}",""&i&"")
			  body_str = GetContent(GetFileText(url),"Flash��Ʒ�б�","��",0)
				response.write "<p>"&url&"����ʼ��ȡ����...</p>" &vbcrlf
			response.Flush
			  body_str_str = split(body_str,"view_b1.gif")
				 for next_i=0 to ubound(body_str_str)
			   next_str = ""&body_str_str(next_i)&""
			   
			   flash_url = "http://flash.86516.com/" &GetContent(next_str,"<a href=""..","""",0) 
			   
			   flash_name = GetContent(next_str,"<font color=""#FE0166"">","</font>",0)
			   if flash_name = "" then
					exit for
			   end if	
			   
					flash_zuozhe = "δ֪"
			  
			   flash_pic = "http://flash.86516.com/" & GetContent(next_str,"border=""0"" src=""../","""",0) 
			   if flash_pic = "" then
					exit for
			   end if 
			 
					   'response.write get_next_url
			   flash_swf = "http://flash.86516.com/" &GetContent(GetFileText(flash_url),"<param name=""movie"" value=""../","""",0)
				if flash_swf = "" then
					exit for
			   end if 
			   
			   flash_jiesao = GetContent(GetFileText(flash_url),"<td width=""80%"" valign=""top"" class=""p1"">","<script language=javascript src=/js/page-rd.js>",0)
			   if flash_jiesao = "" then
				  flash_jiesao = "����"
			   end if

		   


		   set rs=server.createobject("adodb.recordset")  
		  sql="select * from flash where name ='"&flash_name&"'"
			rs.open sql,conn,1,3

			if rs.eof and rs.bof then
			  rs.addnew
			  rs("s_id") = "0" '�����ֶ�
			  rs("username") = "�����"   '������
			  rs("name") = flash_name    'flash��
			  rs("spic") = flash_pic  'ͼƬ��ַ
			  rs("flashswf") = flash_swf    'flash��ַ
			  rs("remark") = flash_jiesao 'flash����
			  rs.update
			Response.Write "��"&next_i&"�� ���ƣ�"&flash_name&"" &vbcrlf
			response.Flush
			Response.Write "  <font color=red> ���</font><br>" &vbcrlf
			response.Flush 


				else
					   response.write "��"&next_i&"����"&flash_name&"<font color=red>�����Ѿ�����</font><br>"  &vbcrlf
				   response.Flush
				end if
				rs.close

			  
			
			Set Rs = Nothing  
		next
		next
			 response.write "<br><font color=red>�ɼ����</font>" &vbcrlf
			 response.Flush



	End Function


	Public Function GetFileText(url) 
     'on error resume next '�д���ʱ����ִ�д���
     Dim http '�������
     'Set http=Server.createobject(XmlHttpCom) '������� 
           Set http=Server.createobject("Microsoft.XMLHTTP") '���������д��һ��������һ�㶼֧�ֵİ汾 
     Http.open "GET",url,False   '�򿪶��� ��GET��ʽ �ȴ���������Ӧ
     Http.Send() '����
     If Http.readystate<>4 Then  '���������û��Ӧ,���˳�����
           Exit Function 
     End If 

     GetFileText=bytes2BSTR(Http.responseBody,"GB2312") 

     Set http=Nothing 
     If err.number<>0 Then err.Clear   '����д���,�������
    End Function

     Function Bytes2bStr(vin,cSet)
       Dim BytesStream,StringReturn
       Set BytesStream = Server.CreateObject("ADODB.Stream")
             BytesStream.Type = 2
             BytesStream.Open
             BytesStream.WriteText vin
             BytesStream.Position = 0
             BytesStream.CharSet = cSet
             BytesStream.Position = 2
             StringReturn =BytesStream.ReadText
             BytesStream.close
              Set BytesStream = Nothing
             Bytes2bStr = StringReturn
     End Function

Public Function GetContent(byref str,byref start,byref last,byref n)

If Instr(lcase(str),lcase(start))>0 then
		select case n
		case 0	
		GetContent=Right(str,Len(str)-Instr(lcase(str),lcase(start))-Len(start)+1)
		GetContent=Left(GetContent,Instr(lcase(GetContent),lcase(last))-1)
		case 1	
		GetContent=Right(str,Len(str)-Instr(lcase(str),lcase(start))+1)
		GetContent=Left(GetContent,Instr(lcase(GetContent),lcase(last))+Len(last)-1)
		case 2	
		GetContent=Right(str,Len(str)-Instr(lcase(str),lcase(start))-Len(start)+1)
		case 3	
		GetContent=Right(str,Len(str)-Instr(lcase(str),lcase(start))+1)
		case 4	
		GetContent=Left(str,InstrRev(lcase(str),lcase(start))+Len(start)-1)
		case 5	
		GetContent=Left(str,InstrRev(lcase(str),lcase(start))-1)
		case 6	
		GetContent=Left(str,Instr(lcase(str),lcase(start))+Len(start)-1)
		case 7	
		GetContent=Right(str,Len(str)-InstrRev(lcase(str),lcase(start))+1)
		case 8	
		GetContent=Left(str,Instr(lcase(str),lcase(start))-1)
		case 9	
		GetContent=Right(str,Len(str)-InstrRev(lcase(str),lcase(start)))
		end select
	Else
		GetContent=""
	End if
End function



End Class
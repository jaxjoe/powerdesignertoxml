'@author:xingxiangyang
'@descrption:输出pd的物理模型为格式化的xml格式
'xml格式：<model><table><c></c></table></model>
'
'
'
Option Explicit
ValidationMode = True
InteractiveMode = im_Abort
dim xml
dim xmlt
dim sort
dim mdl ' 定义当前的模型

'通过全局参数获得当前的模型
Set mdl = ActiveModel
If (mdl Is Nothing) Then
   MsgBox "没有选择模型，请选择一个模型并打开."
 ElseIf Not mdl.IsKindOf(PdPDM.cls_Model) Then
   MsgBox "当前选择的不是一个物理模型（PDM）."
Else
'设置表序号
dim index
index=0
   Dim ForReading, ForWriting, ForAppending
ForReading   = 1 ' Open a file for reading only. You can't write to this file.
ForWriting   = 2 ' Open a file for writing.
ForAppending = 8 ' Open a file and write to the end of the file.
Dim system, file

xml="<?xml version=""1.0"" encoding=""GBK""?><model>"
ProcessFolder mdl


Set system = CreateObject("Scripting.FileSystemObject")
Set file = system.OpenTextFile("d:\tables.xml", ForWriting, true)
xml=""+xml+"</model>"
file.Write xml

 End If


 '--------------------------------------------------------------------------------
 '功能函数
 '--------------------------------------------------------------------------------
Private Sub ProcessFolder(folder)
	dim pack
    dim Tab '定义数据表对象
    dim dsa
     for each pack in folder.AllDiagrams
		output "-"+pack.name
		if   pack.Symbols.Count>0 then
			xml=xml+"<module name='"+ pack.name+"'>"
			for each dsa in pack.Symbols

				if "TableSymbol" =dsa.ObjectType then

					output "       "+dsa.name+"   "+ dsa.ObjectType
					for each tab in folder.tables
						if tab.name=dsa.name then

							output "                             "+tab.name+"   "+ tab.code
							dealTable pack.name,tab
						end if
					next


				end if


			next

			xml=xml+"</module>"
		end if
     next


	dim subfolder
    for each subfolder in folder.Packages
       ProcessFolder subfolder

    next

 End Sub

 private sub dealTable(mcode,tab)


		xmlt=""
		if not tab.isShortcut then
			index=index+1

			xmlt=xmlt+"<table p='"+ mcode+"'  index='"&index&"' code="""+tab.code+""" name="""+tab.name+""">"&chr(13)

          Dim col '定义列对象
          for each col in tab.columns
            ' output col.Domain.DataType

              xmlt=xmlt+"<c code="""+col.code+""" name="""+col.name+""""
              if col.Primary  then
                xmlt=xmlt+" primary=""true"""
              end if
              if col.Domain  Is  Nothing then

               xmlt=xmlt+" Domain="""+col.name+""" "
                xmlt=xmlt+" DomainDataType="""+col.DataType+""" "
             else
               xmlt=xmlt+" Domain="""+col.Domain.name+""" "
                xmlt=xmlt+" DomainDataType="""+col.Domain.DataType+""" "

              end if

              xmlt=xmlt+">"
              xmlt=xmlt+"</c>"&chr(13)
           next
           xmlt=xmlt+"</table>"&chr(13)
      end if
       xml=xml+xmlt


 end sub

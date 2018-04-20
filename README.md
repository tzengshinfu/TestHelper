**必要元件:**  
1.Delphi 6以上

**功能:**  
1.使用DUnit在進行GUI測試時, 截取MessageBox的文字內容。
使用方法如下:

type  
&ensp;&ensp;TTestHelper1 = class(TTestHelper);

var  
&ensp;&ensp;TestHelper1: TTestHelper1;
  
implementation

procedure TTestcase1.SetUp;  
begin  
&ensp;&ensp;TestHelper1 := TTestHelper1.Create();  
end;

procedure TTestcase1.Test_Add;  
var  
&ensp;&ensp;messageBoxContent: string;  
&ensp;&ensp;result: string;  
begin  
&ensp;&ensp;Application.Initialize();  
&ensp;&ensp;Application.CreateForm(TForm1, Form1);  
&ensp;&ensp;Form_Calculater.Show();  

&ensp;&ensp;Form_Calculater.Summand.Text := '1';  
&ensp;&ensp;Form_Calculater.Addend.Text := '2';  
&ensp;&ensp;result := '3';  

&ensp;&ensp;TestHelper1.WaitForMessageBox('Calculater');  
&ensp;&ensp;Form1.Button1.Click;  
&ensp;&ensp;messageBoxContent := TestHelper1.GetMessageBoxContent();  
&ensp;&ensp;Check(messageBoxContent = result, '相加功能異常!!');  
end;
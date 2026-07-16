using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using System.Windows.Forms;
using System.Reflection;
using Microsoft.Vbe.Interop;
using VBIDE = Microsoft.Vbe.Interop;

namespace source
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Excel.Application excelApp = Globals.ThisAddIn.Application;
                Excel.Worksheet ws1 = excelApp.ActiveSheet as Excel.Worksheet;
                if (ws1 == null) return;

                // 获取工作簿
                Excel.Workbook workbook = ws1.Parent as Excel.Workbook;
                if (workbook == null)
                {
                    MessageBox.Show("无法获取工作簿对象");
                    return;
                }

                // 检查是否启用 VBA 访问
                if (!IsVBAccessEnabled(excelApp))
                {
                    MessageBox.Show("请先启用 VBA 访问权限：\n文件 > 选项 > 信任中心 > 信任中心设置 > 宏设置 > 勾选\"信任对 VBA 工程对象模型的访问\"");
                    return;
                }

                // 获取 VBA 工程
                VBIDE.VBProject vbaProject = workbook.VBProject;

                // 获取目标工作表模块
                string componentName = ws1.CodeName;
                VBIDE.VBComponent component = vbaProject.VBComponents.Item(componentName);

                // 创建 SelectionChange 事件代码
                string vbaCode = @"
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    Range(""al6:aw100"").Interior.ColorIndex = xlNone
    If Target.Count > 1 Then
        Set Target = Target.Cells(1)
    End If
    
    If Application.Intersect(Target, Range(""al6:aw100"")) Is Nothing Or Target.Value = """" Then
        Exit Sub
    End If
    
    Dim rng As Range
    For Each rng In Range(""al6:aw100"")
        If rng.Value = Target.Value Then
            rng.Interior.ColorIndex = 36
        End If
    Next rng
End Sub";

                // 删除现有事件（如果存在）
                DeleteExistingEventHandler(component, "Worksheet_SelectionChange");

                // 添加新代码
                component.CodeModule.AddFromString(vbaCode);

                MessageBox.Show("VBA 代码已成功写入工作表！");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"操作失败: {ex.Message}\n\n请确保:\n1. 工作簿已保存为启用宏的格式(.xlsm)\n2. 已启用 VBA 访问权限");
            }
        }

        // 检查是否启用了 VBA 访问权限
        private bool IsVBAccessEnabled(Excel.Application excelApp)
        {
            try
            {
                // 尝试访问 VBProject 属性来检查权限
                var test = excelApp.VBE;
                return true;
            }
            catch
            {
                return false;
            }
        }

        // 删除现有事件处理程序
        private void DeleteExistingEventHandler(VBIDE.VBComponent component, string eventName)
        {
            VBIDE.CodeModule codeModule = component.CodeModule;
            int lineNum = 1;

            while (lineNum <= codeModule.CountOfLines)
            {
                string line = codeModule.get_Lines(lineNum, 1);
                if (line.Contains($"Private Sub {eventName}") ||
                    line.Contains($"Public Sub {eventName}"))
                {
                    int endLine = FindEndSubLine(codeModule, lineNum);
                    codeModule.DeleteLines(lineNum, endLine - lineNum + 1);
                    return;
                }
                lineNum++;
            }
        }

        // 查找 End Sub 行
        private int FindEndSubLine(VBIDE.CodeModule codeModule, int startLine)
        {
            int endLine = startLine;
            int subCount = 1; // 因为我们从找到的Sub开始，所以初始为1

            while (endLine <= codeModule.CountOfLines && subCount > 0)
            {
                endLine++;
                string line = codeModule.get_Lines(endLine, 1).Trim();

                // 检查是否开始新的Sub/Function
                if (line.StartsWith("Sub ", StringComparison.OrdinalIgnoreCase) ||
                    line.StartsWith("Function ", StringComparison.OrdinalIgnoreCase) ||
                    line.StartsWith("Private Sub ", StringComparison.OrdinalIgnoreCase) ||
                    line.StartsWith("Public Sub ", StringComparison.OrdinalIgnoreCase) ||
                    line.StartsWith("Private Function ", StringComparison.OrdinalIgnoreCase) ||
                    line.StartsWith("Public Function ", StringComparison.OrdinalIgnoreCase))
                {
                    subCount++;
                }

                // 检查是否结束Sub/Function
                if (line.Equals("End Sub", StringComparison.OrdinalIgnoreCase) ||
                    line.Equals("End Function", StringComparison.OrdinalIgnoreCase))
                {
                    subCount--;
                }
            }
            return endLine;
        }
    }
}

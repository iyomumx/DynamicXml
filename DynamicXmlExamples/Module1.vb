Imports DynamicXml

Module Module1

    Sub Main()
        '构建XElement对象
        Dim prepare =
            <Settings>
                <File>
                    <FilePath>D:\out\hosts.txt</FilePath>
                    <ReplaceRules>
                        <Search>AAAA</Search>
                        <Replace>BBBB</Replace>
                    </ReplaceRules>
                    <ReplaceRules>
                        <Search>CCCCC0</Search>
                        <Replace>DDDD00</Replace>
                    </ReplaceRules>
                    <RepalceRules>
                        <Search>R18</Search>
                        <Replace>BanIt</Replace>
                    </RepalceRules>
                </File>
                <File>
                    <FilePath>D:\put.lrc</FilePath>
                    <ReplaceRules>
                        <Search>smile0</Search>
                        <Replace>GoG0Go</Replace>
                    </ReplaceRules>
                    <ReplaceRules>
                        <Search>Tea</Search>
                        <Replace>E0</Replace>
                    </ReplaceRules>
                    <ReplaceRules>
                        <Search>10010</Search>
                        <Replace>99990</Replace>
                    </ReplaceRules>
                </File>
            </Settings>
        Dim dn As Object = New XDynamic(prepare)                            '定义为Object以启用后期绑定

        Dim s = From file In CTypeDynamic(Of Object())(dn.File).AsParallel
                            From rr In CTypeDynamic(Of Object())(file.ReplaceRules)
                            Where CStr(DirectCast(rr.Search, XDynamic)).EndsWith("0") AndAlso
                                  CStr(DirectCast(rr.Replace, XDynamic)).Contains("0")
                Select New XDynamic(DirectCast(rr, XDynamic))
        '必须转换为可枚举才能启用查询，必须为Object类型才能启用后期绑定，这里强制使用了CType重载，也可以使用CTypeDynamic
        s.ForAll(Sub(p As Object) p.Good = True)        '使用PLINQ添加项
        For Each r As Object In s
            Console.WriteLine(r.ToString())
        Next
        Console.ReadLine()
        Console.Write(dn.ToString())
        Console.ReadLine()
    End Sub

End Module

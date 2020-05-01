Imports System.Data
Imports SiU.Process.Common.Functions.CommonMethod
Imports SiU.Process.Common.Functions.SystemValue

Public Class DataHelper

    Shared Function CustomizeData(sourceData As DataEntity, targetInfo As TargetInfo) As DataEntity
        Dim rs = sourceData.Clone
        CustomizeMvc(rs, targetInfo)
        CustomizeProxy(rs, targetInfo)
        CustomizeProcess(rs, targetInfo)
        Return rs
    End Function

    Shared Sub CustomizeMvc(sourceData As DataEntity, targetInfo As TargetInfo)
        With sourceData.MvcData
            .Tables(TableName.Programs).
                Select().
                    ToList.ForEach(Function(r)
                                       r.Item("Id") = targetInfo.ProgramCode
                                       r.Item("DisplayName") = targetInfo.ProgramNameKanji
                                       r.Item("Alias") = targetInfo.ProgramCode
                                       r.Item("CreateDateTime") = targetInfo.AdDate
                                       r.Item("CreateUserId") = targetInfo.AdUser
                                       r.Item("UpdateDateTime") = targetInfo.UdDate
                                       r.Item("UpdateUserId") = targetInfo.AdUser
                                       Return True
                                   End Function
                    )
            Dim oldName As String = .Tables(TableName.Programs).Rows(0).Item("DisplayName")
            .Tables(TableName.ToiawaseSearchPanels).
              Select().
                  ToList.ForEach(Function(r)
                                     r.Item("ProgramId") = targetInfo.ProgramCode
                                     r.Item("CreateDateTime") = targetInfo.AdDate
                                     r.Item("CreateUserId") = targetInfo.AdUser
                                     r.Item("UpdateDateTime") = targetInfo.UdDate
                                     r.Item("UpdateUserId") = targetInfo.AdUser
                                     Return True
                                 End Function
                  )

            .Tables(TableName.ToiawaseProperties).
              Select().
                  ToList.ForEach(Function(r)
                                     Dim guid = System.Guid.NewGuid()
                                     r.Item("ProgramId") = targetInfo.ProgramCode
                                     r.Item("Id") = guid
                                     r.Item("Name") = r.Item("Name").ToString.Replace(oldName, targetInfo.ProgramNameKanji)
                                     r.Item("CreateDateTime") = targetInfo.AdDate
                                     r.Item("CreateUserId") = targetInfo.AdUser
                                     r.Item("UpdateDateTime") = targetInfo.UdDate
                                     r.Item("UpdateUserId") = targetInfo.AdUser

                                     Return True
                                 End Function
                  )

            .Tables(TableName.ToiawaseProperties).
            Select("", "ProgramId,UserId,Code").
                ToList.ForEach(Function(r)
                                   Dim xWhere As String = String.Format("UserId= {0} AND Code= {1} ",
                                          fDtSqlStr(r.Item("UserId"), BrowsefSqlStr.fSqlStr_Text),
                                   fDtSqlStr(r.Item("Code"), BrowsefSqlStr.fSqlStr_Integer))
                                   .Tables(TableName.ToiawaseColumnProperties).
                                        Select(xWhere).
                                            ToList.ForEach(Function(r2)
                                                               r2.Item("ProgramId") = targetInfo.ProgramCode
                                                               r2.Item("Id") = r.Item("Id")
                                                               r2.Item("CreateDateTime") = targetInfo.AdDate
                                                               r2.Item("CreateUserId") = targetInfo.AdUser
                                                               r2.Item("UpdateDateTime") = targetInfo.UdDate
                                                               r2.Item("UpdateUserId") = targetInfo.AdUser
                                                               Return True
                                                           End Function
                                            )
                                   .Tables(TableName.ToiawaseReportProperties).
                                      Select(xWhere).
                                          ToList.ForEach(Function(r3)
                                                             r3.Item("ProgramId") = targetInfo.ProgramCode
                                                             r3.Item("Id") = r.Item("Id")
                                                             r3.Item("CreateDateTime") = targetInfo.AdDate
                                                             r3.Item("CreateUserId") = targetInfo.AdUser
                                                             r3.Item("UpdateDateTime") = targetInfo.UdDate
                                                             r3.Item("UpdateUserId") = targetInfo.AdUser

                                                             Return True
                                                         End Function
                                          )

                                   Return True
                               End Function
                )

            .Tables(TableName.ToiawaseDefinations).
              Select().
                  ToList.ForEach(Function(r)
                                     r.Item("ProgramId") = targetInfo.ProgramCode
                                     r.Item("CreateDateTime") = targetInfo.AdDate
                                     r.Item("CreateUserId") = targetInfo.AdUser
                                     r.Item("UpdateDateTime") = targetInfo.UdDate
                                     r.Item("UpdateUserId") = targetInfo.AdUser
                                     Return True
                                 End Function
                  )

            ''TODO
            .Tables(TableName.Menu).
              Select().
                  ToList.ForEach(Function(r)
                                     r.Item("ProgramId") = targetInfo.ProgramCode
                                     'r.Item("Id") = targetInfo.Guid
                                     r.Item("Id") = GetNewMenuId(r.Item("ParentId"), targetInfo.ProgramCode)
                                     r.Item("OrderNumber") = _orderNumber
                                     r.Item("CreateDateTime") = targetInfo.AdDate
                                     r.Item("CreateUserId") = targetInfo.AdUser
                                     r.Item("UpdateDateTime") = targetInfo.UdDate
                                     r.Item("UpdateUserId") = targetInfo.AdUser
                                     Return True
                                 End Function
                  )
        End With

    End Sub

    Private Shared _orderNumber As Integer = 0
    Shared Sub CustomizeProxy(sourceData As DataEntity, targetInfo As TargetInfo)
        With sourceData.ProxyData
            .Tables(TableName.M_PROGRAM_EXTENSION).
                Select().
                    ToList.ForEach(Function(r)
                                       r.Item("Id") = targetInfo.ProgramCode
                                       r.Item("CODE") = targetInfo.ProgramCode
                                       r.Item("HINT_TEXT") = targetInfo.ProgramNameKanji
                                       r.Item("DOCNO") = "DOCNO-" & targetInfo.ProgramCode
                                       r.Item("UD_DATE") = targetInfo.UdDate
                                       r.Item("AD_DATE") = targetInfo.AdDate
                                       r.Item("OPID") = targetInfo.UdOPID
                                       r.Item("UD_USER") = targetInfo.UdUser
                                       Return True
                                   End Function
                    )

            ''TODO
            .Tables(TableName.T_MENU_EXTENSION).
                Select().
                    ToList.ForEach(Function(r)
                                       r.Item("MENU_C") = fToInt(r.Item("MENU_C")) + 10
                                       r.Item("Id") = GetNewMenuId(fToInt(r.Item("MENU_A")), fToInt(r.Item("MENU_B")), targetInfo.ProgramCode)
                                       r.Item("UD_DATE") = targetInfo.UdDate
                                       r.Item("AD_DATE") = targetInfo.AdDate
                                       r.Item("OPID") = targetInfo.UdOPID
                                       r.Item("UD_USER") = targetInfo.UdUser
                                       r.Item("PROGRAM_CODE") = targetInfo.ProgramCode
                                       Return True
                                   End Function
                    )
            '' T_PROGRAM_LINK
            .Tables(TableName.T_PROGRAM_LINK).
                Select().
                ToList.ForEach(Function(r)
                                   r.Item("CODE") = targetInfo.ProgramCode
                                   r.Item("UD_DATE") = targetInfo.UdDate
                                   r.Item("AD_DATE") = targetInfo.AdDate
                                   r.Item("OPID") = targetInfo.UdOPID
                                   r.Item("UD_USER") = targetInfo.UdUser
                                   Return True
                               End Function
                              )

        End With
    End Sub
    Shared Sub CustomizeProcess(sourceData As DataEntity, targetInfo As TargetInfo)
        With sourceData.ProcessData
            .Tables(TableName.M_PROGRAM_ALLOW).
                Select().
                    ToList.ForEach(Function(r)
                                       r.Item("PROGRAM") = targetInfo.ProgramCode
                                       r.Item("UD_DATE") = targetInfo.UdDate
                                       r.Item("AD_DATE") = targetInfo.AdDate
                                       r.Item("OPID") = targetInfo.UdOPID
                                       r.Item("UD_USER") = targetInfo.UdUser
                                       Return True
                                   End Function
                    )
        End With
    End Sub

    Shared Function GetNewMenuId(parentId As String, programCode As String) As String
        Dim dt As DataTable = DbHelper.GetNewMenuId(parentId, programCode)
        _orderNumber = fToInt(dt.Rows(0).Item("OrderNumber")) + 1
        Return GetNewMenuId(dt)

    End Function

    Shared Function GetNewMenuId(menuA As Integer, menuB As Integer, programCode As String) As String
        Dim dt As DataTable = DbHelper.GetNewMenuId(menuA, menuB, programCode)

        Return GetNewMenuId(dt)
    End Function

    Shared Function GetNewMenuId(dt As DataTable) As String

        Dim menuId As String = String.Empty

        If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
            Dim menuC As Integer = fToInt(dt.Rows(0).Item("MENU_C")) + 10
            menuId = "00-" & Right("000" & dt.Rows(0).Item("MENU_A"), 3) & "-" & Right("000" & dt.Rows(0).Item("MENU_B"), 3) & "-" & Right("000" & menuC, 3)

        End If

        Return menuId
    End Function


End Class

/* SELECT * FROM OPCH T0 */
Declare @fdate datetime
declare @Tdate datetime

set @fdate=/* T0.docdate */ '[%0]'
set @Tdate=/* T0.docdate */ '[%1]'

Select A.[Base DocNum],A.[Base DocEntry],A.[Base DocDate],A.[Transaction Type],A.[JE TransId],A.U_JEDoc,
Case when A.[JE TransId] is null then 'N' Else 'Y' End [Flag],A.[Generated Date],A.[Error Code],A.[Error Desc],
Case when A.[JE TransId] is null then 'Failure' Else 'Success' End [Status],A.Status [Log Status]
from 
(Select T0.DocNum [Base DocNum] ,T0.DocEntry [Base DocEntry],T0.DocDate [Base DocDate], 'GRPO' [Transaction Type],T0.U_JEDoc,
Case when T1.U_JETransId is null then T0.U_JEDoc Else T1.U_JETransId End [JE TransId], 
isnull(T1.U_Flag,'N') [Flag],T1.U_GenDate [Generated Date],T1.U_ErrId [Error Code],T1.U_ErrDesc [Error Desc],T1.U_Status [Status]
from OPDN T0 Left Join [@ATPL_SJE] T1 On T0.DocEntry=T1.U_BaseEntry 
Where T0.DocType='S' and T0.CANCELED='N' 
Union all
Select T0.DocNum  ,T0.DocEntry ,T0.DocDate , 'A/P Invoice',T0.U_JEDoc,Case when T1.U_JETransId is null then T0.U_JEDoc Else T1.U_JETransId End,isnull(T1.U_Flag,'N') [Flag],T1.U_GenDate ,T1.U_ErrId ,T1.U_ErrDesc ,T1.U_Status 
from OPCH T0 Left Join [@ATPL_SJE] T1 On T0.DocEntry=T1.U_BaseEntry 
Where T0.DocType='S' and T0.CANCELED='N'  ) A 
Where A.[Base DocDate]>=@fdate and  A.[Base DocDate]<=@Tdate
Order by A.[Base DocDate] desc
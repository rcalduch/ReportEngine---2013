Option Strict On

Public Class SQLBuilder
  Private mSelect As New Text.StringBuilder
  Private mFrom As New Text.StringBuilder
  Private mWhere As New Text.StringBuilder
  Private mOrderBy As New Text.StringBuilder
  Private mGroupBy As New Text.StringBuilder

  ReadOnly Property SQL() As String
    Get
      Dim mSQL As Text.StringBuilder
      mSQL.Length = 0
      If mSelect.Length > 0 Then
        mSQL.Append("SELECT ")
        mSQL.Append(mSelect)
      Else
        mSQL.Append("SELECT * ")
      End If
      If mFrom.Length > 0 Then
        mSQL.AppendFormat("{0} ", mFrom)
      End If
      If mWhere.Length > 0 Then
        mSQL.AppendFormat("{0} ", mWhere)
      End If
      If mOrderBy.Length > 0 Then
        mSQL.Append(mOrderBy)
      End If
      Return mSQL.ToString
    End Get
  End Property

  ReadOnly Property Where() As String
    Get
      If mWhere.Length > 0 Then mWhere.Insert(0, "WHERE ")
      Return mWhere.ToString
    End Get
  End Property

  ReadOnly Property OrderBy() As String
    Get
      If mOrderBy.Length > 0 Then mOrderBy.Insert(0, "ORDER BY ")
      Return mOrderBy.ToString
    End Get
  End Property

  ReadOnly Property WhereOrderBy() As String
    Get
      Return String.Format("{0} {1}", mWhere, mOrderBy)
    End Get
  End Property

  Public Sub AddField(ByVal Field As String)
    If mSelect.Length > 0 Then mSelect.Append(", ")
    mSelect.Append(Field)
  End Sub

  Public Sub AddFrom(ByVal From As String)
    mFrom.AppendFormat("{0} ", From)
  End Sub

  Public Sub AddOrderBy(ByVal Field As String)
    If mOrderBy.Length > 0 Then mOrderBy.Append(", ")
    mOrderBy.Append(Field)
  End Sub

  Public Sub AddWhere(ByVal Field As String, ByVal [Operator] As String, ByVal Value As Object)
    Dim Exp As String
    If Value Is Nothing Then
      Return
    End If
    If mWhere.Length > 0 Then mWhere.Append(" and ")
    Select Case Value.GetType.ToString
      Case "String"
        Exp = CStr(Value)
        If Exp.Length = 0 Then
          Return
        End If
        Select Case [Operator]
          Case "Like"
            Exp = "'" + Exp + "%'"
          Case "LIKE"
            Exp = "'%" + Exp + "%'"
          Case Else
            Exp = "'" + Exp + "'"
        End Select
      Case "Date"
        Exp = String.Format("#{0:d}#", Value)
      Case Else
        Exp = CStr(Value)
    End Select
    mWhere.AppendFormat("(0} {1} {2})", Field, [Operator], Exp)
  End Sub

End Class

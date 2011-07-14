Imports System.Collections


Public Class HL7Helper

    Private SEGMENT_SEPARATOR As String = Chr(13)
    Private FIELD_SEPARATOR As String = "|"
    Private COMPONENT_SEPARATOR As String = "^"
    Private REPETITION_SEPARATOR As String = "~"
    Private ESCAPE_CHARACTER As String = "\"
    Private SUBCOMPONENT_SEPARATOR As String = "&"
    Private NULL As String = """"

    Private MLLP_PREFIX As String = Chr(11)
    Private MLLP_SUFFIX As String = Chr(28) + Chr(13)



    Public Function cleanup(ByVal msgdata As String) As String

        msgdata = msgdata.Replace(MLLP_PREFIX, "")
        msgdata = msgdata.Replace(MLLP_SUFFIX, "")

        Return msgdata

    End Function


    Public Function isValid(ByVal msgdata As String) As Boolean

        If msgdata.Length < 4 Then
            Return False
        ElseIf Not msgdata.StartsWith("MSH") Then
            Return False
        End If

        Return True

    End Function


    ' Retrieve a specified value from an HL7 message
    '
    ' Example: 
    ' Dim HL7 as new HL7
    ' Dim patientNo As String = HL7.getValue(message, "PID", 3, 1)
    '
    Public Function getValue(ByVal msgdata As String, _
                             ByVal segmentName As String, _
                             Optional ByVal fieldId As Integer = 0, _
                             Optional ByVal componentId As Integer = 0, _
                             Optional ByVal subcomponentId As Integer = 0) As String

        Dim segments As String() = Split(msgdata, SEGMENT_SEPARATOR)
        ' Loop through all segments
        For Each segment In segments
            ' If current segment is specified segment
            If segment.Substring(0, 3) = segmentName Then
                ' Check if a field id is specified
                If Not fieldId = Nothing Then
                    ' A field ID was specified so find the specified field
                    Dim fields As String() = Split(segment, FIELD_SEPARATOR)
                    ' Has a component ID been specified?
                    If Not componentId = Nothing Then
                        Dim components As String() = Split(fields(fieldId), COMPONENT_SEPARATOR)
                        If Not subcomponentId = Nothing Then
                            Dim subcomponents As String() = Split(components(componentId - 1), SUBCOMPONENT_SEPARATOR)
                            Return subcomponents(subcomponentId - 1)
                        Else
                            Return components(componentId - 1)
                        End If
                    Else
                        ' No component ID was specified, just return entire field
                        Return fields(fieldId - 1)
                    End If
                Else
                    ' No field ID so return entire segment
                    Return segment
                End If
            End If
        Next

    End Function


    Public Property empty() As String
        Get
            Return NULL
        End Get
        Set(ByVal value As String)
            NULL = value
        End Set
    End Property


    Public Property seg_sep() As String
        Get
            Return SEGMENT_SEPARATOR
        End Get
        Set(ByVal value As String)
            SEGMENT_SEPARATOR = value
        End Set
    End Property


    Public Property fld_sep() As String
        Get
            Return FIELD_SEPARATOR
        End Get
        Set(ByVal value As String)
            FIELD_SEPARATOR = value
        End Set
    End Property


    Public Property cmp_sep() As String
        Get
            Return COMPONENT_SEPARATOR
        End Get
        Set(ByVal value As String)
            COMPONENT_SEPARATOR = cmp_sep
        End Set
    End Property


    Public Property rep_sep() As String
        Get
            Return REPETITION_SEPARATOR
        End Get
        Set(ByVal value As String)
            REPETITION_SEPARATOR = value
        End Set
    End Property


    Public Property esc_chr() As String
        Get
            Return ESCAPE_CHARACTER
        End Get
        Set(ByVal value As String)
            ESCAPE_CHARACTER = value
        End Set
    End Property


    Public Property sub_sep() As String
        Get
            Return SUBCOMPONENT_SEPARATOR
        End Get
        Set(ByVal value As String)
            SUBCOMPONENT_SEPARATOR = value
        End Set
    End Property


    Public Property llp_pfx() As String
        Get
            Return MLLP_PREFIX
        End Get
        Set(ByVal value As String)
            MLLP_PREFIX = value
        End Set
    End Property


    Public Property llp_sfx() As String
        Get
            Return MLLP_SUFFIX
        End Get
        Set(ByVal value As String)
            MLLP_SUFFIX = value
        End Set
    End Property


End Class

Attribute VB_Name = "CalendarUtility"
'<License>------------------------------------------------------------
'
' Copyright (c) 2020 Shinnosuke Yakenohara
'
' This program is free software: you can redistribute it and/or modify
' it under the terms of the GNU General Public License as published by
' the Free Software Foundation, either version 3 of the License, or
' (at your option) any later version.
'
' This program is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
' GNU General Public License for more details.
'
' You should have received a copy of the GNU General Public License
' along with this program.  If not, see <http://www.gnu.org/licenses/>.
'
'-----------------------------------------------------------</License>

'
' �J�����_�[��̎w�������N�Z���āA�w�莞�Ԃ�����I�������Z�o���ĕԂ�
'
' Parameters
' ----------
' startDay : Date
'   �N�Z�J�n��
'   �������܂܂��ꍇ�́A���̎�������N�Z�J�n����
'
' hours : Double
'   ����鎞�Ԑ��̍��v[h]
'
' calendar : Range or String
'   �����̏���Ԑ���`�\�B�ȉ��̏����𖞂�����`�ł��邱��
'    - 1��ڂ� Date �^�̓��t�A2��ڂɂ��̓��ɏ���鎞�Ԑ�[h](Double�^�ɃL���X�g�\�Ȑ��l)����`����Ă��邱��
'    - 1�s����1������`����Ă���A���t�̏ȗ��s�����݂��Ȃ�����(e.g. 1��/1���̎��̍s��1��/3���ł����Ă͂Ȃ�Ȃ�
'
' targetDay : Date, Default Nothing
'   �w�肳�ꂽ�ꍇ�A���̓��܂łɏ���鎞�Ԑ�[h] ��Ԃ�
'
' integral : Boolean, Default True
'   targetDay���w�肳�ꂽ�ꍇ�ɂ̂ݗL��
'   True �� �w�肳�ꂽ�ꍇ�AtargetDay �̓��܂łɏ���鎞�Ԑ�[h] ��Ԃ�
'   False �� �w�肳�ꂽ�ꍇ�AtargetDay �̓��ɏ���鎞�Ԑ�[h] ��Ԃ�
'
Public Function calendarHowSpend( _
    ByVal startDay As Date, _
    ByVal hours As Double, _
    ByVal calendar As Variant, _
    Optional ByVal targetDay As Variant = Nothing, _
    Optional ByVal integral As Boolean = True _
) As Variant

    Dim date_startDay As Date
    Dim date_endDay As Date
    Dim date_targetDay As Variant
    Dim dbl_startHour As Double
    Dim long_rowOfStartDay As Long
    Dim vararr_calendar As Variant
    
    Dim int_dateCol As Integer
    Dim int_hourCol As Integer
    
    Dim expendedHoursPerDay As Double
    Dim expendedHours As Double
    Dim long_spendDays As Long
    Dim dbl_remainInDay As Double
    
    Dim toRet As Variant
    
    date_startDay = (Year(startDay) & "/" & Month(startDay) & "/" & Day(startDay)) '���Ԃ�j�����ē��t���擾
    dbl_startHour = (hour(startDay) + Minute(startDay) / 60)                       '���t��j�����Ď��Ԃ��擾

    Set date_targetDay = Nothing
    If TypeName(targetDay) = "Range" Then ' Range �^�̏ꍇ
    
        If TypeName(targetDay.Value) <> "Date" Then ' Range ���\���Z�����̒l�̌^�� Date �^�ł͂Ȃ��ꍇ
            toRet = xlErrValue ' #VALUE! ��Ԃ�
            GoTo RETURN_AND_EXIT
            
        End If
        
        date_targetDay = targetDay.Value
        
    ElseIf TypeName(targetDay) = "Date" Then ' Date �^�̏ꍇ
        date_targetDay = targetDay
        
    End If
    
    str_tmp = TypeName(date_targetDay)
    If (str_tmp <> "Date" And str_tmp <> "Nothing") Then
        toRet = xlErrValue ' #VALUE! ��Ԃ�
        GoTo RETURN_AND_EXIT

    End If
    
    If TypeName(date_targetDay) = "Date" Then
        If (date_targetDay < date_startDay) Then
            toRet = 0 ' 0 ��Ԃ�
            GoTo RETURN_AND_EXIT
        End If
    End If
    
    bool_foundStartDay = False
    
    If TypeName(calendar) = "String" Then ' �����̏���Ԑ���`�\�� String �^�Ŏw�肳�ꂽ�ꍇ
        vararr_calendar = Range(calendar)
        
    Else ' �����̏���Ԑ���`�\�� String �^�ȊO�̏ꍇ
        vararr_calendar = calendar ' Range �I�u�W�F�N�g�Ƃ݂Ȃ�
        
    End If
    
    
    '�J�n�����J�����_�[���猟������
    For long_rowOfStartDay = LBound(vararr_calendar, 1) To UBound(vararr_calendar, 1)
        
        int_dateCol = LBound(vararr_calendar, 2) '��ԍ��̗����t��Ƃ݂Ȃ�
        int_hourCol = int_dateCol + 1            '��ԍ� + 1 �̗�����Ԓ�`��Ƃ݂Ȃ�
        
        If (vararr_calendar(long_rowOfStartDay, int_dateCol) = date_startDay) Then '�J�n��������������
            bool_foundStartDay = True
            Exit For
        End If
        
    Next long_rowOfStartDay
    
    If (Not bool_foundStartDay) Then '�J�n����������Ȃ�������
        toRet = xlErrValue ' #VALUE! ��Ԃ�
        
    Else '�J�n��������������
        
        expendedHoursPerDay = 0 '���̓��ɏ���鎞�Ԑ�
        expendedHours = 0 '�ώZ�����
        long_spendDays = 0  '�������(0 based)
        
        
        '�^�X�N�I�������w����܂œ������Ԑ���ώZ���郋�[�v
        bl_isFirst = True
        Do While True
            
            If bl_isFirst Then ' ���[�v�̈��ڂ̏ꍇ
                
                '�����̎c�莞��(����\����)�̎Z�o
                dbl_remainInDay = (vararr_calendar(long_rowOfStartDay + long_spendDays, int_hourCol) * 1) - dbl_startHour
                
                ' �ŏ��̓��̊J�n���Ԃ����̓��̎��Ԑ��𒴂��Ă���ꍇ
                ' e.g. �ŏ��̓��̊J�n���Ԃ� 9:00 �����ǁA���̓��̎��Ԑ��� 8.00 [h]
                ' �������A0 [h] �͋��e����B���̓�����ώZ�J�n���鎖�Ɠ��`������B
                ' e.g. �ŏ��̓��̊J�n���Ԃ� 8:00 �����ǁA���̓��̎��Ԑ��� 8.00 [h]
                '      -> ����͎��̓��� 0:00 ����ώZ�J�n���鎖�Ɠ����B
                If dbl_remainInDay < 0 Then
                    toRet = xlErrValue ' #VALUE! ��Ԃ�
                    GoTo RETURN_AND_EXIT
                End If
                
                bl_isFirst = False
                
            Else
                '�J�����_�[�� 1�� ��s�ƂȂ��Ă��邩�ǂ����m�F
                int_expectedAs1Day = DateDiff("d", vararr_calendar(long_rowOfStartDay + long_spendDays - 1, int_dateCol), vararr_calendar(long_rowOfStartDay + long_spendDays, int_dateCol))
                If (int_expectedAs1Day <> 1) Then '�J�����_�[�� 1�� ��s�ƂȂ��Ă��Ȃ��ꍇ
                    toRet = xlErrValue ' #VALUE! ��Ԃ�
                    GoTo RETURN_AND_EXIT
                End If
                
                '�����̎c�莞��(����\����)�̎Z�o
                dbl_remainInDay = (vararr_calendar(long_rowOfStartDay + long_spendDays, int_hourCol) * 1)
            
            End If
            
            bool_tmp = False
            If TypeName(date_targetDay) = "Date" Then
                bool_tmp = (vararr_calendar(long_rowOfStartDay + long_spendDays, int_dateCol) = date_targetDay)
            End If
            
            
            '���̓��Ń^�X�N�̎��Ԃ����ׂď���I���ꍇ
            If (hours <= (expendedHours + dbl_remainInDay)) Then
                
                long_flacHour = 0
                If long_spendDays = 0 Then ' ���[�v�̈��ڂ̏ꍇ
                    long_flacHour = dbl_startHour '�Z�o����I�����Ԃɂ��炩���ߊJ�n���Ԃ𑫂��Ă���
                End If
                
                long_flacHour = long_flacHour + (hours - expendedHours) '�I�����Ԃ̎Z�o
                expendedHoursPerDay = hours - expendedHours
                expendedHours = hours '�ώZ����Ԃ̓^�X�N�̎��Ԑ��ɓ�����
                
                Exit Do
                
            ElseIf bool_tmp Then '�w����̏ꍇ
                
                long_flacHour = 0
                If long_spendDays = 0 Then ' ���[�v�̈��ڂ̏ꍇ
                    long_flacHour = dbl_startHour '�Z�o����I�����Ԃɂ��炩���ߊJ�n���Ԃ𑫂��Ă���
                End If
                
                long_flacHour = long_flacHour + (hours - expendedHours) '�I�����Ԃ̎Z�o
                expendedHoursPerDay = dbl_remainInDay
                expendedHours = expendedHours + dbl_remainInDay '�ώZ����Ԃ̓^�X�N�̎��Ԑ��ɓ�����
                
                Exit Do
                
            '���̓��Ń^�X�N�̎��Ԃ����ׂď���I���Ȃ� And
            '�w����ł��Ȃ��ꍇ
            Else
                expendedHoursPerDay = dbl_remainInDay
                expendedHours = expendedHours + dbl_remainInDay '�ώZ����Ԃɂ��̓��̏���\���Ԃ�a�Z
                long_spendDays = long_spendDays + 1 '�������(0 based) += 1
                
            End If
            
        Loop
        
        If TypeName(date_targetDay) = "Date" Then
            
            If integral Then '�ώZ����Ԏw��̏ꍇ
                toRet = expendedHours
                
            Else  '�w����̏���Ԏw��̏ꍇ
            
                date_endDay = vararr_calendar(long_rowOfStartDay + long_spendDays, int_dateCol)
                
                If date_endDay < date_targetDay Then
                    toRet = 0
                    
                Else
                    toRet = expendedHoursPerDay
                    
                End If
                
                'todo �J�n���������
                
            End If

        Else
            date_endDay = vararr_calendar(long_rowOfStartDay + long_spendDays, int_dateCol)
            date_endDay = DateAdd("h", long_flacHour, date_endDay)
            date_endDay = DateAdd("n", ((long_flacHour - Fix(long_flacHour)) * 60), date_endDay)
            toRet = date_endDay

        End If
        
        
    End If
    
RETURN_AND_EXIT:
    
    calendarHowSpend = toRet
    
End Function

'
' �J�����_�[��̎w���������N�Z���āA�������Ԍ�̓�����Ԃ�
' e.g. 1��1���̏���\����[h] �� 8[h] �ƃJ�����_�[�ɒ�`����Ă��āA
'      �w������� "1/1 8:00" �������Ƃ��́A"1/2 0:00" ��Ԃ�
'
' Parameters
' ----------
' dateToAdd : Date
'   �N�Z�J�n��
'   �������܂܂��ꍇ�́A���̎�������N�Z�J�n����
'
' calendar : Range or String
'   �����̏���Ԑ���`�\�B�ȉ��̏����𖞂�����`�ł��邱��
'    - 1��ڂ� Date �^�̓��t�A2��ڂɂ��̓��ɏ���鎞�Ԑ�[h](Double�^�ɃL���X�g�\�Ȑ��l)����`����Ă��邱��
'    - 1�s����1������`����Ă���A���t�̏ȗ��s�����݂��Ȃ�����(e.g. 1��/1���̎��̍s��1��/3���ł����Ă͂Ȃ�Ȃ�
'
Public Function calendarAddDelta( _
    ByVal dateToAdd As Date, _
    ByVal calendar As Variant _
) As Variant
    
    Dim date_startDay As Date
    Dim dbl_startHour As Double
    Dim long_rowOfStartDay As Long
    Dim vararr_calendar As Variant
    
    Dim int_dateCol As Integer
    Dim int_hourCol As Integer

    Dim long_spendDays As Long
    Dim dbl_remainInDay As Double
    
    date_startDay = (Year(dateToAdd) & "/" & Month(dateToAdd) & "/" & Day(dateToAdd)) '���Ԃ�j�����ē��t���擾
    dbl_startHour = (hour(dateToAdd) + Minute(dateToAdd) / 60)                        '���t��j�����Ď��Ԃ��擾
    
    bool_foundStartDay = False
    
    If TypeName(calendar) = "String" Then ' �����̏���Ԑ���`�\�� String �^�Ŏw�肳�ꂽ�ꍇ
        vararr_calendar = Range(calendar)
        
    Else ' �����̏���Ԑ���`�\�� String �^�ȊO�̏ꍇ
        vararr_calendar = calendar ' Range �I�u�W�F�N�g�Ƃ݂Ȃ�
        
    End If
    
    '�N�Z�J�n�����J�����_�[���猟������
    For long_rowOfStartDay = LBound(vararr_calendar, 1) To UBound(vararr_calendar, 1)
        
        int_dateCol = LBound(vararr_calendar, 2) '��ԍ��̗����t��Ƃ݂Ȃ�
        int_hourCol = int_dateCol + 1            '��ԍ� + 1 �̗�����Ԓ�`��Ƃ݂Ȃ�
        
        If (vararr_calendar(long_rowOfStartDay, int_dateCol) = date_startDay) Then '�N�Z�J�n��������������
            bool_foundStartDay = True
            Exit For
        End If
        
    Next long_rowOfStartDay
    
    If (Not bool_foundStartDay) Then '�N�Z�J�n����������Ȃ�������
        toRet = xlErrValue ' #VALUE! ��Ԃ�
        
    Else '�N�Z�J�n��������������
        
        dbl_remainInDay = vararr_calendar(long_rowOfStartDay, int_hourCol) - dbl_startHour

        ' �N�Z�J�n���̎��Ԃ����̓��̎��Ԑ��𒴂��Ă���ꍇ
        ' e.g. �N�Z�J�n���̊J�n���Ԃ� 9:00 �����ǁA���̓��̎��Ԑ��� 8.00 [h]
        If dbl_remainInDay < 0 Then
            Err.Raise (xlErrValue)  ' �����I�� #VALUE! ��Ԃ�

        ElseIf dbl_remainInDay = 0 Then ' �N�Z�J�n���̎��Ԃ����̓��̎��Ԑ��ƈ�v����ꍇ

            '���̊J�n������������
            long_spendDays = 0
            Do While True

                long_spendDays = long_spendDays + 1
                
                If 0 < (vararr_calendar(long_rowOfStartDay + long_spendDays, int_hourCol) * 1) Then
                    toRet = vararr_calendar(long_rowOfStartDay + long_spendDays, int_dateCol)
                    Exit Do
                End If

            Loop

        Else  ' �N�Z�J�n���̎��Ԃ����̓��̎��Ԑ���菭�Ȃ��ꍇ
            toRet = dateToAdd

        End If
        
    End If
    
    calendarAddDelta = toRet
    
End Function


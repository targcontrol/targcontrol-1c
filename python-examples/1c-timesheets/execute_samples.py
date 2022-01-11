# общая часть запроса загрузки месяца из 28 дней в 1С
def timesheet_load_sample(cursor, conn, id_table, id1c, hours):
    cursor.execute(f"UPDATE _Document427_VT16125  "
                   f"SET _Fld16128 = {hours[0]}, _Fld16129 = {hours[1]}, "
                   f"_Fld16130 = {hours[2]}, _Fld16131 = {hours[3]}, _Fld16132 = {hours[4]},"
                   f" _Fld16133 = {hours[5]},"
                   f" _Fld16134 = {hours[6]}, _Fld16135 = {hours[7]}, _Fld16136 = {hours[8]}, "
                   f"_Fld16137 = {hours[9]}, _Fld16138 = {hours[10]}, _Fld16139 = {hours[11]}, "
                   f"_Fld16140 = {hours[12]}, _Fld16141 = {hours[13]}, _Fld16142 = {hours[14]},"
                   f" _Fld16143 = {hours[15]}, _Fld16144 = {hours[16]}, _Fld16145 = {hours[17]}, _Fld16146 = {hours[18]},"
                   f" _Fld16147 = {hours[19]}, "
                   f"_Fld16148 = {hours[20]}, _Fld16149 = {hours[21]}, _Fld16150 = {hours[22]}, _Fld16151 = {hours[23]}, "
                   f"_Fld16152 = {hours[24]}, _Fld16153 = {hours[25]}, _Fld16154 = {hours[26]}, _Fld16155 = {hours[27]}"
                   f" where _Document427_IDRRef = {id_table} and _Fld16127RRef = {id1c}")
    #conn.commit()


def calendarevents_load_sample(cursor, conn, id_table, id1c, events):
    cursor.execute(f"UPDATE _Document427_VT16125  "
                   f"SET _Fld16159RRef = {events[0]}, _Fld16160RRef = {events[1]}, "
                   f"_Fld16161RRef = {events[2]}, _Fld16162RRef = {events[3]}, _Fld16163RRef = {events[4]},"
                   f" _Fld16164RRef = {events[5]},"
                   f" _Fld16165RRef = {events[6]}, _Fld16166RRef = {events[7]}, _Fld16167RRef = {events[8]}, "
                   f"_Fld16168RRef = {events[9]}, _Fld16169RRef = {events[10]}, _Fld16170RRef = {events[11]}, "
                   f"_Fld16171RRef = {events[12]}, _Fld16172RRef = {events[13]}, _Fld16173RRef = {events[14]},"
                   f" _Fld16174RRef = {events[15]}, _Fld16175RRef = {events[16]}, _Fld16176RRef = {events[17]}, _Fld16177RRef = {events[18]},"
                   f" _Fld16178RRef = {events[19]}, "
                   f"_Fld16179RRef = {events[20]}, _Fld16180RRef = {events[21]}, _Fld16181RRef = {events[22]}, _Fld16182RRef = {events[23]}, "
                   f"_Fld16183RRef = {events[24]}, _Fld16184RRef = {events[25]}, _Fld16185RRef = {events[26]}, _Fld16186RRef = {events[27]}"
                   f" where _Document427_IDRRef = {id_table} and _Fld16127RRef = {id1c}")
    #conn.commit()
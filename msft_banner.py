# banner items
wellcome_str = 'Üdvözöl a  Máv-Start Forda Ellenőrző program!'

# main_version.calendar.daily_relase
version = 'ver: 0.200103.2'

# rows
full_row = ['{0:#^55}', '\n', ]
blind_row = ['{0:<}#', '{0:>53}#', '\n']
str1_row = ['{0:<}#', '{1:^53}', '#', '\n']
str2_row = ['{0:<}#', '{2:^53}', '#', '\n']

# this is the ready banner
banner = [full_row, blind_row, str1_row,
          blind_row, str2_row, blind_row, full_row]

# Done
done = True


def make_banner(word_list):
    """ Elkészíti az üdvözlő szöveget."""
    wellcome_msg = ''
    for i in word_list:
        if type(i) is list:
            # list unpacking
            for k in i:
                wellcome_msg = wellcome_msg + k
        else:
            wellcome_msg = wellcome_msg + i

    print(wellcome_msg.format('', wellcome_str, version))
    print('\n')

    return done

#!/usr/bin/python
# ParagonySDS - program do pobierania raportów dobowych i zliczania płatności kartą oraz gotówką
# Autor: Piotr Rozwadowski rozwadowski@sds.uw.edu.pl siotsonek@gmail.com

# aktualizacje:
# 1.0.1 - 10.02.2020 - dodano obsługę reszt przy płatnościach
# 2.0.2 - 11.02.2020 - poprawki kosmetyczne, dodanie [ENTER] i focus na dacie

import imaplib
import re
import datetime
import tkinter as tk
from tkinter import messagebox

VERSION = "1.0.2"
AUTHOR = "rozwadowski@sds.uw.edu.pl"
CONF_FILE = "mail.conf"
TAXE_RATE = {"A": 0.23, "B": 0.08, "C": 0.05, "D": 0.0, "E": 0.0}


def pick_mails(date, results):
    mail_index = []
    for res in results:
        regex = "[0-9]{2}-[0-9]{2}-[0-9]{4}"
        if len(re.findall(regex, res)) > 0:
            all_dates = re.findall(regex, res)
            dates = list(set(all_dates))
            dates_cnt = [all_dates.count(i) for i in dates]
            max_dates = dates_cnt.index(max(dates_cnt))
            if date == dates[max_dates]:
                mail_index.append(results.index(res))
    return mail_index


def processing(mail_index, results, payment):
    grs, tax, cnt = 0., {"A": 0., "B": 0., "C": 0., "D": 0., "E": 0.}, 0
    net = {"A": 0., "B": 0., "C": 0., "D": 0., "E": 0.}
    for mail_ind in mail_index:
        lines = results[mail_ind].split("\n")
        i = 0
        for line in lines:
            regex = payment + ":[ ]+([0-9.]+)"
            if len(re.findall(regex, line)) > 0:
                grs += float(re.findall(regex, line)[0])
                change_re = 'Reszta \(' + payment + ' PLN\):[ ]+([0-9.]+)'
                if len(re.findall(change_re, lines[i + 1])) > 0:
                    grs -= float(re.findall(change_re, lines[i + 1])[0])
                cnt += 1

                start_line = i
                end_line = start_line
                while len(re.findall("RAZEM:", lines[start_line])) == 0:
                    start_line = start_line - 1

                for k in range(start_line + 2, end_line):
                    for t in TAXE_RATE.keys():
                        taxed_re = "Sprzedaż opodatkowana " + t + ":[ ]+([0-9.]+)"
                        free_re = "Sprzedaż zwolniona " + t + ":[ ]+([0-9.]+)"
                        if len(re.findall(taxed_re, lines[k])) > 0:
                            sp = float(re.findall(taxed_re, lines[k])[0])
                            net[t] += float(sp / (1.0 + TAXE_RATE[t]))
                            tax[t] += sp - sp / (TAXE_RATE[t] + 1.0)
                        if len(re.findall(free_re, lines[k])) > 0:
                            sp = float(re.findall(free_re, lines[k])[0])
                            net[t] += float(sp / (1.0 + TAXE_RATE[t]))
            i += 1
    # return gross, netto, tax, number of receipts, netto in rates
    return grs, round(grs - sum([round(t, 2) for t in tax.values()]), 2), [round(t, 2) for t in tax.values()], cnt, \
           [round(t, 2) for t in net.values()]  # netto in rates


def summary(mail_index, results):
    grs, net, tax, cnt = 0., 0., 0., 0
    for i in mail_index:
        lines = results[i].split("\n")
        for line in lines:
            regex = "Należność:[ ]+PLN ([0-9.]+)"
            if len(re.findall(regex, line)) > 0:
                grs += float(re.findall(regex, line)[0])
                cnt += 1
                regex = "SUMA PTU[ ]+([0-9.]+)"
                tax += float(re.findall(regex, lines[lines.index(line) - 2])[0])
            regex = "Sprzedaż netto w stawce [ABCD]{1}[ ]+([0-9.]+)"
            if len(re.findall(regex, line)) > 0:
                net += float(re.findall(regex, line)[0])
            regex = "Sprzedaż zwolniona E[ ]+([0-9.]+)"
            if len(re.findall(regex, line)) > 0:
                net += float(re.findall(regex, line)[0])
    return round(grs, 2), round(net, 2), round(tax, 2), cnt


def raport(window, imap, login, psw, date, mail_from, folder, reg_name):
    # grab data from entries
    imap_server = imap.get()
    mail_login = login.get()
    mail_pass = psw.get()
    mail_date = [int(i) for i in date.get().split("-")]
    mail_from = mail_from.get()
    mail_folder = folder.get()
    cash_name = reg_name.get()
    # parse date
    date_since = (datetime.date(mail_date[2], mail_date[1], mail_date[0]) - datetime.timedelta(5)).strftime("%d-%b-%Y")
    date_before = (datetime.date(mail_date[2], mail_date[1], mail_date[0]) + datetime.timedelta(5)).strftime("%d-%b-%Y")
    # connect to server
    try:
        mail = imaplib.IMAP4_SSL(imap_server)
        mail.login(mail_login, mail_pass)
        mail.list()
        mail.select(mail_folder)
        # search email
        status, response = mail.search(None, '(FROM "' + mail_from + '")', 'SUBJECT "*[' + cash_name + ']*"',
                                       '(SINCE {date})'.format(date=date_since),
                                       '(BEFORE {date})'.format(date=date_before))
        messages_uid = re.findall("[0-9]+", str(response[0]))

        results = []
        for uid in messages_uid:
            _, res = mail.fetch(uid, '(UID BODY[TEXT])')
            results.append(res[0][1].decode("utf-8"))
    except (imaplib.IMAP4.error, TimeoutError, imaplib.socket.gaierror):
        messagebox.showinfo("Błąd", "Nieudana próba połączenia z pocztą")
        results = []

    mail_index = pick_mails(date.get(), results)
    # processing mails
    cash_grs, cash_net, cash_tax, cnt_cash, cash_netABCDE = processing(mail_index, results, "Gotówka")
    card_grs, card_net, card_tax, cnt_card, card_netABCDE = processing(mail_index, results, "Karta")
    charge_grs, charge_net, charge_tax, cnt_charge = summary(mail_index, results)

    # RAPORT CREATING
    # left column
    rap1 = "Raporty:\nPłatności kartą:\nPłatności gotówką:\n" + "".ljust(20, "-") + "\nKarta \n"
    for i in range(len(card_tax)):
        rap1 += "PTU " + "ABCDE"[i] + ":\t" + str(card_tax[i]) + "\n"
    rap1 += "Netto: \nBrutto: \n" + "".ljust(20, "-") + "\nGotówka \n"
    for i in range(len(cash_tax)):
        rap1 += "PTU " + "ABCDE"[i] + ":\t" + str(cash_tax[i]) + "\n"
    rap1 += "Netto: \nBrutto: \n" + "".ljust(20, "-") + "\nPodsumowanie\nPTU \nNetto: \nBrutto:"
    # right column
    rap2 = str(len(mail_index)) + "\n" + str(cnt_card) + "\n" + str(cnt_cash) + "\n" \
           + "".ljust(20, "-") + "\n" + "\n"  # KARTA
    for i in range(len(card_tax)):
        rap2 += "Netto " + "ABCDE"[i] + "\t" + str(card_netABCDE[i]) + "\n"
    rap2 += str(card_net) + "\n" + str(card_grs) + "\n" + "".ljust(20, "-") + "\n\n"
    for i in range(len(cash_tax)):  # GOTÓWKA
        rap2 += "Netto " + "ABCDE"[i] + "\t" + str(cash_netABCDE[i]) + "\n"
    rap2 += str(cash_net) + "\n" + str(cash_grs) + "\n" + "".ljust(20, "-") + "\n\n"  # PODSUMOWANIE
    rap2 += str(charge_tax) + "\n" + str(charge_net) + "\n" + str(charge_grs) + "\n"

    blackboard1 = tk.Text(window, height=27, width=20)
    blackboard1.grid(row=9, column=0)
    blackboard1.insert(tk.END, rap1)
    blackboard2 = tk.Text(window, height=27, width=20)
    blackboard2.grid(row=9, column=1)
    blackboard2.insert(tk.END, rap2)

    # WARNINGS
    if abs(cash_grs + card_grs - charge_grs) > 10e-3:
        messagebox.showinfo("Błąd", "Niezgodne kwoty brutto")
    if abs(cash_net + card_net - charge_net) > 10e-3:
        messagebox.showinfo("Błąd", "Niezgodne kwoty netto")
    if abs(sum(cash_netABCDE) - cash_net) > 10e-3:
        messagebox.showinfo("Błąd", "Niezgodne kwoty netto według stawek A-E (gotówka)")
    if abs(sum(card_netABCDE) - card_net) > 10e-3:
        messagebox.showinfo("Błąd", "Niezgodne kwoty netto według stawek A-E (karta)")
    if abs(sum(cash_tax) + sum(card_tax) - charge_tax) > 10e-3:
        messagebox.showinfo("Błąd", "Niezgodna suma PTU")
    if len(mail_index) > 1:
        messagebox.showinfo("Uwaga", "Pobrano więcej niż jeden raport dobowy")
    if len(mail_index) != cnt_charge:
        messagebox.showinfo("Błąd", "Niezgodna liczba raportów")


def main():
    window = tk.Tk()
    window.title("Paragony SDS v." + VERSION)
    window.geometry("380x720")

    try:
        file = open(CONF_FILE)
        imap_server = file.readline()[:-1]
        mail_login = file.readline()[:-1]
        mail_pass = file.readline()[:-1]
        mail_date = datetime.date.today().strftime("%d-%m-%Y")
        mail_from = file.readline()[:-1]
        mail_folder = file.readline()[:-1]
        cash_name = file.readline()[:-1]
    except FileNotFoundError:
        messagebox.showinfo("Błąd", "Nie znaleziono pliku konfiguracyjnego")
        imap_server = mail_login = mail_pass = ""
        mail_date = datetime.date.today().strftime("%d-%m-%Y")
        mail_from = mail_folder = cash_name = ""

    tk.Label(window, text=AUTHOR).grid(row=0, column=0)
    labels = ["<--- Kontakt", "Serwer imap", "Login", "Hasło", "Data", "Mail źródłowy", "Folder", "Kasa fiskalna"]
    for label in labels:
        tk.Label(window, text=label).grid(row=labels.index(label), column=1)

    imapEntry = tk.Entry(window, width=24)
    loginEntry = tk.Entry(window, width=24)
    passEntry = tk.Entry(window, width=24, show="●")
    dateEntry = tk.Entry(window, width=24)
    dateEntry.focus_force()
    fromEntry = tk.Entry(window, width=24)
    folderEntry = tk.Entry(window, width=24)
    regNameEntry = tk.Entry(window, width=24)

    entries = [imapEntry, loginEntry, passEntry, dateEntry, fromEntry, folderEntry, regNameEntry]
    inserts = [imap_server, mail_login, mail_pass, mail_date, mail_from, mail_folder, cash_name]

    for entry in entries:
        entry.grid(row=entries.index(entry) + 1, column=0)
        entry.insert(0, inserts[entries.index(entry)])

    raportBtn = tk.Button(window, text="Raport", command=lambda: raport(window, imapEntry, loginEntry, passEntry,
                                                                        dateEntry, fromEntry, folderEntry,
                                                                        regNameEntry))
    raportBtn.grid(row=8, column=0)
    window.bind('<Return>', (lambda event: raport(window, imapEntry, loginEntry, passEntry, dateEntry, fromEntry,
                                                  folderEntry, regNameEntry)))
    window.mainloop()


if __name__ == "__main__":
    main()

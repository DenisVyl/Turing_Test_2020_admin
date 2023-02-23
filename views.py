import csv
import xlwt
from django.shortcuts import render, get_object_or_404
from django.http import HttpResponse, JsonResponse
from django.contrib.auth.decorators import login_required
from django import forms
from .models import Chats
from .models import Messages
from .models import Subjects
from .models import Testers


def home(request):
    return render(request, 'adminka/home.html')  # HttpResponse('<h1>Welcome to the Turing Test home page!<h1>')


@login_required
def export_xls(request):
    response = HttpResponse(content_type='application/ms-excel')
    response['Content-Disposition'] = 'attachment; filename="All_data.xls"'

    workbook = xlwt.Workbook(encoding='utf-8')
    ws_alldata = workbook.add_sheet('All_data')
    ws_alldata.col(0).width = 1_000
    ws_alldata.col(1).width = 10_000
    ws_alldata.col(5).width = 5_000
    ws_alldata.col(6).width = 5_000
    ws_alldata.col(7).width = 5_000

    row_num = 0

    font_style = xlwt.XFStyle()
    font_style.font.bold = True
    ws_alldata.write(row_num, 0, 'Chats', font_style)
    row_num += 1

    columns = ['id', 'public_id', 'tester', 'subject', 'status', 'answer']

    for col_num in range(len(columns)):
        ws_alldata.write(row_num, col_num, columns[col_num], font_style)

    font_style = xlwt.XFStyle()

    rows = Chats.objects.all().values_list('id', 'public_id', 'tester', 'subject', 'status', 'answer')
    for row in range(len(rows)):
        row_num += 1
        for col in range(len(rows[0])):
            ws_alldata.write(row_num, col, str(rows[row][col]), font_style)

    row_num += 2
    font_style = xlwt.XFStyle()
    font_style.font.bold = True
    ws_alldata.write(row_num, 0, 'Messages', font_style)
    row_num += 1
    columns = ['id', 'public_id', 'chat', 'subject', 'tester', 'time', 'request_text', 'response_text', 'context']

    for col_num in range(len(columns)):
        ws_alldata.write(row_num, col_num, columns[col_num], font_style)

    font_style = xlwt.XFStyle()

    rows = Messages.objects.all().values_list('id', 'public_id', 'chat', 'subject', 'tester', 'time',
                                              'request_text', 'response_text', 'context')
    for row in range(len(rows)):
        row_num += 1
        for col in range(len(rows[0])):
            ws_alldata.write(row_num, col, str(rows[row][col]), font_style)

    row_num += 2
    font_style = xlwt.XFStyle()
    font_style.font.bold = True
    ws_alldata.write(row_num, 0, 'Subjects', font_style)
    row_num += 1
    columns = ['id', 'public_id', 'type', 'name', 'webhook', 'livetex_department_id', 'status']
    for col_num in range(len(columns)):
        ws_alldata.write(row_num, col_num, columns[col_num], font_style)

    font_style = xlwt.XFStyle()

    rows = Subjects.objects.all().values_list('id', 'public_id', 'type', 'name', 'webhook', 'livetex_department_id',
                                              'status')
    for row in range(len(rows)):
        row_num += 1
        for col in range(len(rows[0])):
            ws_alldata.write(row_num, col, str(rows[row][col]), font_style)

    row_num += 2
    font_style = xlwt.XFStyle()
    font_style.font.bold = True
    ws_alldata.write(row_num, 0, 'Subjects', font_style)
    row_num += 1
    columns = ['id', 'public_id', 'telegram_id', 'telegram_chat_id', 'telegram_username', 'telegram_first_name',
               'telegram_last_name', 'status']
    for col_num in range(len(columns)):
        ws_alldata.write(row_num, col_num, columns[col_num], font_style)

    font_style = xlwt.XFStyle()

    rows = Testers.objects.all().values_list('id', 'public_id', 'telegram_id', 'telegram_chat_id', 'telegram_username',
                                             'telegram_first_name',
                                             'telegram_last_name', 'status')
    for row in range(len(rows)):
        row_num += 1
        for col in range(len(rows[0])):
            ws_alldata.write(row_num, col, str(rows[row][col]), font_style)

    workbook.save(response)
    return response


@login_required
def export(request):
    """exports all tables from database"""
    response = HttpResponse(content_type='text/csv')
    response.write(u'\ufeff'.encode('utf8'))
    writer = csv.writer(response)

    writer.writerow(['ID', 'Public Id', 'Tester', 'Subject', 'Status', 'Answer'])
    for chat in Chats.objects.all().values_list('id', 'public_id', 'tester', 'subject', 'status', 'answer'):
        writer.writerow(chat)
    writer.writerow([])

    writer.writerow(
        ['id', 'public_id', 'chat', 'subject', 'tester', 'time', 'request_text', 'response_text', 'context'])
    for message in Messages.objects.all().values_list('id', 'public_id', 'chat', 'subject', 'tester', 'time',
                                                      'request_text', 'response_text', 'context'):
        writer.writerow(message)
    writer.writerow([])

    writer.writerow(['id', 'public_id', 'type', 'name', 'webhook', 'livetex_department_id', 'status'])
    for subject in Subjects.objects.all().values_list('id', 'public_id', 'type', 'name', 'webhook',
                                                      'livetex_department_id', 'status'):
        writer.writerow(subject)
    writer.writerow([])

    writer.writerow(['id', 'public_id', 'telegram_id', 'telegram_chat_id', 'telegram_username', 'telegram_first_name',
                     'telegram_last_name', 'status'])
    for tester in Testers.objects.all().values_list('id', 'public_id', 'telegram_id', 'telegram_chat_id',
                                                    'telegram_username', 'telegram_first_name', 'telegram_last_name',
                                                    'status'):
        writer.writerow(tester)
    writer.writerow([])

    response['Content-Disposition'] = 'attachment; filename="All_data.csv"'
    return response


@login_required
def export_dialogs_xls(request):
    response = HttpResponse(content_type='application/ms-excel')
    response['Content-Disposition'] = 'attachment; filename="Dialogs.xls"'

    workbook = xlwt.Workbook(encoding='utf-8')
    ws_dialogs = workbook.add_sheet('Dialogs')

    ws_dialogs.col(0).width = 5_000
    ws_dialogs.col(1).width = 5_000
    ws_dialogs.col(2).width = 5_000
    ws_dialogs.col(3).width = 5_000
    ws_dialogs.col(4).width = 5_000
    ws_dialogs.col(5).width = 5_000

    row_num = 0

    font_style = xlwt.XFStyle()
    font_style.font.bold = True
    ws_dialogs.write(row_num, 0, 'Все сообщения из каждого диалога, отсортированные по времени:', font_style)
    row_num += 1

    columns = ['Диалог:', 'Время:', 'Проверяемый:', 'Тестировщик:', 'Вопрос:', 'Ответ:', ]
    for col_num in range(len(columns)):
        ws_dialogs.write(row_num, col_num, columns[col_num], font_style)

    font_style = xlwt.XFStyle()

    rows = Messages.objects.order_by('chat', 'time').values_list('chat', 'time', 'subject', 'tester', 'request_text',
                                                                 'response_text', )
    for row in range(len(rows)):
        row_num += 1
        for col in range(len(rows[0])):
            ws_dialogs.write(row_num, col, str(rows[row][col]), font_style)

    row_num += 1
    font_style = xlwt.XFStyle()
    font_style.font.bold = True
    ws_dialogs.write(row_num, 0, f'Всего сообщений: {len(list(Messages.objects.all()))}', font_style)
    row_num += 1
    ws_dialogs.write(row_num, 0, f'Всего диалогов: {len(list(Chats.objects.all()))}', font_style)

    workbook.save(response)
    return response


@login_required
def export_dialogs(request):
    """exports all messages from each chat, sort by timestamp"""
    response = HttpResponse(content_type='text/csv')
    response.write(u'\ufeff'.encode('utf8'))
    writer = csv.writer(response)
    total_messages = 0

    writer.writerow(['Все сообщения из каждого диалога, отсортированные по времени:'])
    writer.writerow(['Диалог:', 'Время:', 'Проверяемый:', 'Тестировщик:',
                     'Вопрос:', 'Ответ:', ])
    for message in Messages.objects.order_by('chat', 'time').values_list('chat', 'time', 'subject', 'tester',
                                                                         'request_text', 'response_text', ):
        writer.writerow(message)

    writer.writerow([f'Всего сообщений: {len(list(Messages.objects.all()))}'])
    writer.writerow([f'Всего диалогов: {len(list(Chats.objects.all()))}'])

    response['Content-Disposition'] = 'attachment; filename="Dialogs.csv"'
    return response


@login_required
def export_testers_and_chats_xls(request):
    response = HttpResponse(content_type='application/ms-excel')
    response['Content-Disposition'] = 'attachment; filename="Testers_and_chats.xls"'

    workbook = xlwt.Workbook(encoding='utf-8')
    ws_tnc = workbook.add_sheet('Testers_and_chats')

    ws_tnc.col(0).width = 5_000
    ws_tnc.col(1).width = 5_000

    row_num = 0

    font_style = xlwt.XFStyle()
    font_style.font.bold = True
    ws_tnc.write(row_num, 0,
                 'В таблице представлены пользователи, тестировавшие более 1 раза, и кол-во проведённых диалогов, в порядке убывания',
                 font_style)
    row_num += 1

    columns = ['Имя пользователя:', 'Количество его диалогов:']
    for col_num in range(len(columns)):
        ws_tnc.write(row_num, col_num, columns[col_num], font_style)

    font_style = xlwt.XFStyle()

    d = {}
    for tester in Testers.objects.all():
        d.update({tester.telegram_username: 0})

    for chat in Chats.objects.all():
        d[chat.tester.telegram_username] += 1

    d = {tester: chat for tester, chat in sorted(d.items(), key=lambda item: item[1], reverse=True)}

    for tester_telegram_username in d.keys():
        if d[tester_telegram_username] >= 2:
            row_num += 1
            ws_tnc.write(row_num, 0, tester_telegram_username, font_style)
            ws_tnc.write(row_num, 1, d[tester_telegram_username], font_style)

    workbook.save(response)
    return response


@login_required
def export_testers_and_chats(request):
    """exports users who tested more than once

    группируем чаты по пользователям, считаем для каждого пользователя количество чатов
    выводим в таблицу парами "пользователь - число его чатов". Сортируем по числу чатов по убыванию
    """
    response = HttpResponse(content_type='text/csv')
    response.write(u'\ufeff'.encode('utf8'))
    writer = csv.writer(response)

    writer.writerow([
        'В таблице представлены пользователи, тестировавшие более 1 раза, и кол-во проведённых диалогов, в порядке убывания'])
    writer.writerow(['Имя пользователя:', 'Количество его диалогов:'])
    d = {}
    for tester in Testers.objects.all():
        d.update({tester.telegram_username: 0})

    for chat in Chats.objects.all():
        d[chat.tester.telegram_username] += 1

    d = {tester: chat for tester, chat in sorted(d.items(), key=lambda item: item[1], reverse=True)}

    for tester_telegram_username in d.keys():
        if d[tester_telegram_username] >= 2:
            writer.writerow([tester_telegram_username, d[tester_telegram_username]])

    response['Content-Disposition'] = 'attachment; filename="Testers_and_chats.csv"'
    return response


@login_required
def export_bots_misrecognized_xls(request):
    response = HttpResponse(content_type='application/ms-excel')
    response['Content-Disposition'] = 'attachment; filename="Bots.xls"'

    workbook = xlwt.Workbook(encoding='utf-8')
    ws_bots = workbook.add_sheet('Bots')

    row_num = 0

    font_style = xlwt.XFStyle()
    font_style.font.bold = True

    columns = ['Имя бота:', 'Ошибочно приняли за человека:', 'Правильно опознали как бота:', 'Всего диалогов провёл:']
    for col_num in range(len(columns)):
        ws_bots.write(row_num, col_num, columns[col_num], font_style)

    font_style = xlwt.XFStyle()

    for bot in Subjects.objects.filter(type='bot'):
        count_as_human = 0
        count_as_bot = 0
        number_of_chats = 0
        for chat in Chats.objects.filter(subject=bot, answer='human'):
            count_as_human += 1
        for chat in Chats.objects.filter(subject=bot, answer='bot'):
            count_as_bot += 1
        for chat in Chats.objects.filter(subject=bot):
            number_of_chats += 1
        row_num += 1
        ws_bots.write(row_num, 0, bot.name, font_style)
        ws_bots.write(row_num, 1, count_as_human, font_style)
        ws_bots.write(row_num, 2, count_as_bot, font_style)
        ws_bots.write(row_num, 3, number_of_chats, font_style)

    workbook.save(response)
    return response


@login_required
def export_bots_misrecognized(request):
    """Number of times the bot was voted as a human"""
    response = HttpResponse(content_type='text/csv')
    response.write(u'\ufeff'.encode('utf8'))
    writer = csv.writer(response)

    writer.writerow(
        ['Имя бота:', 'Ошибочно приняли за человека:', 'Правильно опознали как бота:', 'Всего диалогов провёл:'])

    for bot in Subjects.objects.filter(type='bot'):
        count_as_human = 0
        count_as_bot = 0
        number_of_chats = 0
        for chat in Chats.objects.filter(subject=bot, answer='human'):
            count_as_human += 1
        for chat in Chats.objects.filter(subject=bot, answer='bot'):
            count_as_bot += 1
        for chat in Chats.objects.filter(subject=bot):
            number_of_chats += 1
        writer.writerow([bot.name, count_as_human, count_as_bot, number_of_chats])

    response['Content-Disposition'] = 'attachment; filename="Bots.csv"'
    return response


@login_required
def export_humans_misrecognized_xls(request):
    response = HttpResponse(content_type='application/ms-excel')
    response['Content-Disposition'] = 'attachment; filename="Humans.xls"'

    workbook = xlwt.Workbook(encoding='utf-8')
    ws_humans = workbook.add_sheet('Humans')

    row_num = 0

    font_style = xlwt.XFStyle()
    font_style.font.bold = True

    columns = ['Имя волонтёра:', 'Ошибочно приняли за бота:', 'Правильно опознали как человека:',
               'Всего диалогов провёл:']
    for col_num in range(len(columns)):
        ws_humans.write(row_num, col_num, columns[col_num], font_style)

    font_style = xlwt.XFStyle()

    for human in Subjects.objects.filter(type='volunteer'):
        count_as_bot = 0
        count_as_human = 0
        number_of_chats = 0
        for chat in Chats.objects.filter(subject=human, answer='bot'):
            count_as_bot += 1
        for chat in Chats.objects.filter(subject=human, answer='human'):
            count_as_human += 1
        for chat in Chats.objects.filter(subject=human):
            number_of_chats += 1
        row_num += 1
        ws_humans.write(row_num, 0, human.name, font_style)
        ws_humans.write(row_num, 1, count_as_bot, font_style)
        ws_humans.write(row_num, 2, count_as_human, font_style)
        ws_humans.write(row_num, 3, number_of_chats, font_style)

    workbook.save(response)
    return response


@login_required
def export_humans_misrecognized(request):
    """Number of times the human was voted as a bot"""
    response = HttpResponse(content_type='text/csv')
    response.write(u'\ufeff'.encode('utf8'))
    writer = csv.writer(response)

    writer.writerow(
        ['Имя волонтёра:', 'Ошибочно приняли за бота:', 'Правильно опознали как человека:', 'Всего диалогов провёл:'])
    for human in Subjects.objects.filter(type='volunteer'):
        count_as_bot = 0
        count_as_human = 0
        number_of_chats = 0
        for chat in Chats.objects.filter(subject=human, answer='bot'):
            count_as_bot += 1
        for chat in Chats.objects.filter(subject=human, answer='human'):
            count_as_human += 1
        for chat in Chats.objects.filter(subject=human):
            number_of_chats += 1
        writer.writerow([human.name, count_as_bot, count_as_human, number_of_chats])

    response['Content-Disposition'] = 'attachment; filename="Humans.csv"'
    return response


@login_required
def export_statistics(request):
    """statistics for a bot"""
    response = HttpResponse(content_type='text/csv')
    response.write(u'\ufeff'.encode('utf8'))
    writer = csv.writer(response)

    writer.writerow(['name', 'as human', 'as bot', 'number of chats'])

    queryset = Subjects.objects.filter(type='bot')
    for bot in queryset:
        as_human = 0
        as_bot = 0
        for chat in Chats.objects.filter(subject=bot, answer='human'):
            as_human += 1
        for chat in Chats.objects.filter(subject=bot, answer='bot'):
            as_bot += 1
        writer.writerow([as_human, as_bot])

    response['Content-Disposition'] = 'attachment; filename="bots_statistics.csv"'
    return response


def status(request):
    """requesting bot status by uuid"""
    uuid_chosen = request.GET.get('bot_id')

    if uuid_chosen != '' and uuid_chosen is not None:
        q = Subjects.objects.all()
        try:
            subject = get_object_or_404(q, public_id=uuid_chosen)
            context = {"ok": subject.status == "active"}
            if not context["ok"]:
                context.update(({"code": subject.status}))
            response = JsonResponse(context, content_type='application/json')
        except forms.ValidationError:
            response = HttpResponse("Not found", status=404)
    else:
        response = HttpResponse("UUID required", status=400)

    return response

from apscheduler.schedulers.blocking import BlockingScheduler
import update_ivr as ui
import update_opera as uo
import update_kpo as uk
import update_mail as um
import functions.serverfile_update as sfu
import update_richiamate as ur
import update_ibia as uib
import update_ibiasos as uibs


def c87survey2hh():
    ui.update_ivr()
    um.mail_c87()
    uib.update_ibia()


def update_daily():
    uo.update_opera()
    sfu.dfo_update()
    ur.update_richiamate()
    uibs.start_connection()


def update_twoforweek():
    uo.update_kpo()


sched = BlockingScheduler()

sched.add_job(c87survey2hh, "cron", day_of_week="mon-sun", hour="8-21/2", minute=35)
sched.add_job(update_daily, "cron", day_of_week="mon-sun", hour="17", minute=00)
sched.add_job(update_twoforweek, "cron", day_of_week="mon, thu", hour="15", minute=00)
sched.start()

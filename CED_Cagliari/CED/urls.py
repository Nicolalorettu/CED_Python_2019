from django.urls import path
from django.conf.urls import url
from . import views
from django.contrib.auth import views as auth_views

app_name = 'CED'
urlpatterns = [

    path('', views.login_view, name='login'),
    path('', views.logout_view, name='logout'),
    url(r'^login/', views.checklogin, name='check'),
    path('index/', views.index, name='index'),
    path('download/', views.download, name='download'),
    url(r'^ivr/', views.ivr, name='ivr'),
    url(r'^sos_easy_sms/', views.sos_easy_sms, name='sos_easy_sms'),
    url(r'^richiamate/', views.richiamate, name='richiamate'),
    url(r'^opera/', views.opera, name='opera'),
    url(r'^kpi_c87/', views.kpi_c87, name='kpi_c87'),
    url(r'^kpi_ibia/', views.kpi_ibia, name='kpi_ibia'),
    url(r'^Sos_Easy_PAF/', views.sos_easy_paf, name='sos_easy_paf'),
    url(r'^under_construction/', views.under_constr, name='under_constr'),
    url(r'^password_reset/$', auth_views.password_reset, name='password_reset'),
    url(r'^password_reset/done/$', auth_views.password_reset_done, name='password_reset_done'),
    url(r'^reset/(?P<uidb64>[0-9A-Za-z_\-]+)/(?P<token>[0-9A-Za-z]{1,13}-[0-9A-Za-z]{1,20})/$', auth_views.password_reset_confirm, name='password_reset_confirm'),
    url(r'^reset/done/$', auth_views.password_reset_complete, name='password_reset_complete'),

]

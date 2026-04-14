"""URL patterns for the scheduler app (v4)."""

from django.urls import path
from . import views

urlpatterns = [
    path('', views.index, name='index'),
    path('upload/', views.upload_file, name='upload'),
    path('generate/', views.generate, name='generate'),
    path('timetable/', views.timetable_view, name='timetable'),
    path('dashboard/', views.dashboard_view, name='dashboard'),
    path('api/timetable/', views.api_timetable, name='api_timetable'),
    path('api/timetable/edit/', views.api_edit, name='api_edit'),
    path('api/timetable/validate/', views.api_validate, name='api_validate'),
    path('api/timetable/partial-regen/', views.api_partial_regen, name='api_partial_regen'),
    path('api/explain/<str:section>/<str:day>/<str:slot>/', views.api_explain, name='api_explain'),
    path('api/debug-schedule/', views.api_debug_schedule, name='api_debug_schedule'),
    path('api/teacher-load/', views.api_teacher_load, name='api_teacher_load'),
    path('download-pdf/', views.download_pdf, name='download_pdf'),
    path('download-xlsx/', views.download_xlsx, name='download_xlsx'),
]

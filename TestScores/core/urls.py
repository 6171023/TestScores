from django.urls import path
from . import views

urlpatterns = [
    path('', views.upload, name='upload_home'), 
    path('success/', views.success, name='success'),
    path('upload/', views.upload, name='upload'),
    path('download/<str:file_name>/', views.download_file, name='download_file'),
]

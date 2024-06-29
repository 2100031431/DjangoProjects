from django.contrib import admin
from django.urls import path
from myapp1 import views

urlpatterns = [
    path('admin/', admin.site.urls),
    path('', views.index, name='index'),
    path('download/', views.download_excel, name='download_excel'),
]

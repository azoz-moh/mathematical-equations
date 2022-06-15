from django.urls import path
from . import views

urlpatterns = [
    path('', views.index, name='index'),
    path('download-pdf', views.some_view, name="download-pdf"),
    path('pdf-file', views.pdf_file, name="pdf-file"),
    path('excel1/', views.excel1, name="excel1"),
    path('excel2/', views.excel2, name="excel2"),
    path('excel3/', views.excel3, name="excel3"),
    path('f/', views.pdf_view, name="pdf_view"),
]

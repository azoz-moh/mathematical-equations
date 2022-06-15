from django.urls import path
from . import views

urlpatterns = [
    path('', views.index, name='ind'),
    path('ex/', views.ex, name='ex'),
                   ]

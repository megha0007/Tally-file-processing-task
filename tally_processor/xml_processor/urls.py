from django.urls import path
from . import views

urlpatterns = [
    path('process-tally/', views.process_tally, name='process_tally'),
]

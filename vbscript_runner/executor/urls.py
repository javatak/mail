from django.urls import path
from .views import RunVBScript

urlpatterns = [
    path('run-script/', RunVBScript.as_view(), name='run-script'),
]

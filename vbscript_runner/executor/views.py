import subprocess
from rest_framework.views import APIView
from rest_framework.response import Response
from rest_framework import status

class RunVBScript(APIView):
    def post(self, request):
        script_path = "C:\\path\\to\\your\\script.vbs"
        try:
            result = subprocess.run(['cscript', script_path], capture_output=True, text=True)
            return Response({"output": result.stdout, "error": result.stderr}, status=status.HTTP_200_OK)
        except Exception as e:
            return Response({"error": str(e)}, status=status.HTTP_500_INTERNAL_SERVER_ERROR)

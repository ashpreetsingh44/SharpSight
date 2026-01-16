from flask import Flask, render_template, request, send_file
import cv2
import numpy as np
from openpyxl import Workbook
import tempfile

app = Flask(
    __name__,
    template_folder="../templates",
    static_folder="../static"
)


def is_blurry(image_bytes, threshold=100.0):
    npimg = np.frombuffer(image_bytes, np.uint8)
    image = cv2.imdecode(npimg, cv2.IMREAD_COLOR)

    if image is None:
        return True, 0.0

    gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
    lap_var = cv2.Laplacian(gray, cv2.CV_64F).var()

    return lap_var < threshold, lap_var


@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        files = request.files.getlist("images")

        wb = Workbook()
        ws = wb.active
        ws.title = "Blur Results"
        ws.append(["S.No", "Image Name", "Blur", "Laplacian Variance"])

        row = 1

        for f in files:
            img_bytes = f.read()
            blurry, score = is_blurry(img_bytes)

            ws.append([
                row,
                f.filename,
                "Yes" if blurry else "No",
                round(score, 2)
            ])
            row += 1

        tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
        wb.save(tmp.name)
        tmp.close()

        return send_file(tmp.name, as_attachment=True, download_name="blur_results.xlsx")

    return render_template("index.html")


app = app

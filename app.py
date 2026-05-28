from flask import Flask,Response, send_file, request, jsonify
import pandas as pd
import io
import unicodedata

app = Flask(__name__)

# =====================================================
# URL GOOGLE SHEETS CSV
# =====================================================
URL_CSV = (
    "https://docs.google.com/spreadsheets/d/"
    "1-ZE_bRAGQtYq2D6iSd79AYzPxek9KOAbikWEFhW9oYw"
    "/export?format=csv"
)

URL_PUNTOS = (
    "https://docs.google.com/spreadsheets/d/"
    "1-ZE_bRAGQtYq2D6iSd79AYzPxek9KOAbikWEFhW9oYw"
    "/export?format=csv&gid=0"
)
URL_ACNB = (
    "https://docs.google.com/spreadsheets/d/"
    "1-ZE_bRAGQtYq2D6iSd79AYzPxek9KOAbikWEFhW9oYw"
    "/export?format=csv&gid=801021966"
)

def normalizar_texto(texto):

    if pd.isna(texto):
        return ""

    texto = str(texto)

    texto = unicodedata.normalize("NFKD", texto)

    texto = ''.join(
        c for c in texto
        if not unicodedata.combining(c)
    )

    texto = (
        texto
        .replace(".0", "")
        .replace("\xa0", " ")
        .strip()
        .upper()
    )

    texto = " ".join(texto.split())

    return texto

# =====================================================
# ENDPOINT PUNTOS DE PAGO
# =====================================================
@app.route("/descargar/puntos_pago")
def descargar_puntos_pago():

    try:

        # =============================================
        # PARAMETROS
        # =============================================
        ut = request.args.get("ut")
        depa = request.args.get("depa")
        prov = request.args.get("prov")
        dist = request.args.get("dist")

        print("UT:", ut)
        print("DEPA:", depa)
        print("PROV:", prov)
        print("DIST:", dist)

        # =============================================
        # LEER CSV
        # =============================================
        df = pd.read_csv(
            URL_PUNTOS,
            dtype=str
        )
        # =============================================
        # NORMALIZAR TEMPORALMENTE
        # =============================================
        df.columns = (
            df.columns
            .str.strip()
            .str.lower()
        )
        # =============================================
        # NORMALIZAR DATA
        # =============================================
        for col in [
            "ut",
            "departamento",
            "provincia",
            "distrito"
        ]:
            if col in df.columns:
                df[col] = df[col].apply(normalizar_texto)

        # =============================================
        # NORMALIZAR PARAMS
        # =============================================
        ut = normalizar_texto(ut) if ut else None
        depa = normalizar_texto(depa) if depa else None
        prov = normalizar_texto(prov) if prov else None
        dist = normalizar_texto(dist) if dist else None

        print(df["ut"].unique()[:20])
        print("UT RECIBIDO:", repr(ut))

        # =============================================
        # FILTROS
        # =============================================
        if ut and ut != "ALL":
            df = df[df["ut"] == ut]

        if depa and depa != "ALL":
            df = df[df["departamento"] == depa]

        if prov and prov != "ALL":
            df = df[df["provincia"] == prov]

        if dist and dist != "ALL":
            df = df[df["distrito"] == dist]

        # =============================================
        # GENERAR EXCEL
        # =============================================
        columnas_numericas = [
            "USUARIOS_ABONADOS",
            "PUNTOS_PAGO"
        ]
        output = io.BytesIO()

        with pd.ExcelWriter(
            output,
            engine="openpyxl" #"xlsxwriter"

        ) as writer:
            # =============================================
            # RESTAURAR MAYUSCULAS
            # =============================================
            df.columns = [
                col.upper()
                for col in df.columns
            ]
            # Convierte a numeros las columnas numericas
            for col in columnas_numericas:

                if col in df.columns:

                    df[col] = pd.to_numeric(
                        df[col],
                        errors="coerce"
                    )

            df.to_excel(
                writer,
                index=False,
                sheet_name="PuntosPago"
            )

            workbook = writer.book
            # agrega formato numerico
            number_format = workbook.add_format({
                "num_format": "#,##0" # con decimales: "num_format": "#,##0.00"
            })

            worksheet = writer.sheets["PuntosPago"]
            if df.empty:
                worksheet.write("A2", "SIN REGISTROS")

            # =========================================
            # HEADER FORMAT
            # =========================================
            header_format = workbook.add_format({
                "bold": True,
                "font_color": "white",
                "bg_color": "#024270",
                "align": "center",
                "border": 1
            })

            for col_num, value in enumerate(df.columns):

                worksheet.write(
                    0,
                    col_num,
                    value,
                    header_format
                )

            # =========================================
            # AUTO WIDTH
            # =========================================
            for i, col in enumerate(df.columns):

                if df.empty:
                    max_len = len(col)
                else:
                    max_len = max(
                        df[col]
                        .astype(str)
                        .map(len)
                        .max(),
                        len(col)
                    )

                width = min(max_len + 2, 50)

                # columnas numericas
                if col in columnas_numericas:

                    worksheet.set_column(
                        i,
                        i,
                        width,
                        number_format
                    )

                else:

                    worksheet.set_column(
                        i,
                        i,
                        width
                    )

        output.seek(0)

        # =============================================
        # RESPONSE
        # =============================================
        excel_data = output.getvalue()

        return Response(
            excel_data,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={
                "Content-Disposition":
                    "attachment; filename=Reporte_Puntos_Pago.xlsx",
                "Connection": "close",
                "Cache-Control": "no-store"
            }
        )

    except Exception as e:

        return jsonify({
            "status": 500,
            "error": str(e)
        }), 500

# =====================================================
# ENDPOINT ACNB
# =====================================================
@app.route("/descargar/acnb", methods=["GET"])
def descargar_acnb():

    try:
        depa = request.args.get("depa")
        prov = request.args.get("prov")
        dist = request.args.get("dist")
        
        # =============================================
        # NORMALIZAR PARAMS
        # =============================================
        depa = normalizar_texto(depa) if depa else None
        prov = normalizar_texto(prov) if prov else None
        dist = normalizar_texto(dist) if dist else None

        # =============================================
        # LEER CSV
        # =============================================
        df = pd.read_csv(URL_ACNB,dtype=str)

        # =============================================
        # NORMALIZAR COLUMNAS
        # =============================================
        df.columns = df.columns.str.strip().str.lower()

        for col in [
            "departamento",
            "provincia",
            "distrito"
        ]:
            if col in df.columns:
                df[col] = df[col].apply(normalizar_texto)
        # =============================================
        # FILTROS
        # =============================================
        if depa and depa != "ALL":
            df = df[df["departamento"] == depa]

        if prov and prov != "ALL":
            df = df[df["provincia"] == prov]

        if dist and dist != "ALL":
            df = df[df["distrito"] == dist]

        # =============================================
        # GENERAR EXCEL
        # =============================================
        output = io.BytesIO()

        with pd.ExcelWriter(
            output,
            engine="openpyxl" #"xlsxwriter"
        ) as writer:
            # =============================================
            # RESTAURAR MAYUSCULAS
            # =============================================
            df.columns = [
                col.upper()
                for col in df.columns
            ]
            df.to_excel(
                writer,
                index=False,
                sheet_name="ACNB"
            )

            workbook = writer.book

            worksheet = writer.sheets["ACNB"]
            
            if df.empty:
                worksheet.write("A2", "SIN REGISTROS")

            # =========================================
            # HEADER FORMAT
            # =========================================
            header_format = workbook.add_format({
                "bold": True,
                "font_color": "white",
                "bg_color": "#024270",
                "align": "center",
                "border": 1
            })

            for col_num, value in enumerate(df.columns):

                worksheet.write(
                    0,
                    col_num,
                    value,
                    header_format
                )

            # =========================================
            # AUTO WIDTH
            # =========================================
            for i, col in enumerate(df.columns):

                if df.empty:
                    max_len = len(col)
                else:
                    max_len = max(
                        df[col]
                        .astype(str)
                        .map(len)
                        .max(),
                        len(col)
                    )

                worksheet.set_column(
                    i,
                    i,
                    min(max_len + 2, 50)
                )

        output.seek(0)

        # =============================================
        # RESPONSE
        # =============================================
        excel_data = output.getvalue()

        return Response(
            excel_data,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={
                "Content-Disposition":
                    "attachment; filename=Reporte_ACNB.xlsx",
                "Connection": "close",
                "Cache-Control": "no-store"
            }
        )

    except Exception as e:

        return jsonify({
            "status": 500,
            "error": str(e)
        }), 500

# HEALTHCHECK
@app.route("/health")
def health():
    return jsonify({
        "status": "ok"
    }), 200

# =====================================================
# MAIN
# =====================================================
if __name__ == "__main__":
    import os
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)  
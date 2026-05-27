from flask import Flask, send_file, request, jsonify
import pandas as pd
import io

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

            df[col] = (
                df[col]
                .fillna("")
                .astype(str)
                .str.replace(".0", "", regex=False)
                .str.strip()
                .str.upper()
            )

        # =============================================
        # NORMALIZAR PARAMS
        # =============================================
        ut = str(ut).strip().upper() if ut else None
        depa = str(depa).strip().upper() if depa else None
        prov = str(prov).strip().upper() if prov else None
        dist = str(dist).strip().upper() if dist else None

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
        # VALIDAR VACIO
        # =============================================
        if df.empty:

            return jsonify({
                "status": 204,
                "message": "Sin registros"
            }), 204

        # =============================================
        # GENERAR EXCEL
        # =============================================
        output = io.BytesIO()

        with pd.ExcelWriter(
            output,
            engine="xlsxwriter"
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
                sheet_name="PuntosPago"
            )

            workbook = writer.book

            worksheet = writer.sheets["PuntosPago"]

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
        return send_file(
            output,
            as_attachment=True,
            download_name="Reporte_Puntos_Pago.xlsx",
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:

        return jsonify({
            "status": 500,
            "error": str(e)
        }), 500

# =====================================================
# ENDPOINT ACNB
# =====================================================
@app.route("/descargar/acnb")
def descargar_acnb():

    try:

        # =============================================
        # PARAMETROS
        # =============================================
        depa = request.args.get("depa")
        prov = request.args.get("prov")
        dist = request.args.get("dist")

        print("DEPA:", depa)
        print("PROV:", prov)
        print("DIST:", dist)

        # =============================================
        # LEER CSV
        # =============================================
        df = pd.read_csv(URL_ACNB)

        # =============================================
        # NORMALIZAR COLUMNAS
        # =============================================
        df.columns = df.columns.str.strip().str.lower()

        # =============================================
        # FILTROS
        # =============================================
        if depa and depa != "ALL":

            df = df[
                df["departamento"]
                .astype(str)
                .str.upper()
                == depa.upper()
            ]

        if prov and prov != "ALL":

            df = df[
                df["provincia"]
                .astype(str)
                .str.upper()
                == prov.upper()
            ]

        if dist and dist != "ALL":

            df = df[
                df["distrito"]
                .astype(str)
                .str.upper()
                == dist.upper()
            ]

        # =============================================
        # VALIDAR VACIO
        # =============================================
        if df.empty:

            return jsonify({
                "status": 204,
                "message": "Sin registros"
            }), 204

        # =============================================
        # GENERAR EXCEL
        # =============================================
        output = io.BytesIO()

        with pd.ExcelWriter(
            output,
            engine="xlsxwriter"
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
        return send_file(
            output,
            as_attachment=True,
            download_name="Reporte_ACNB.xlsx",
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:

        return jsonify({
            "status": 500,
            "error": str(e)
        }), 500


# =====================================================
# MAIN
# =====================================================
if __name__ == "__main__":
    import os
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)  
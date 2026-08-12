from flask import Flask, render_template, request, send_file, flash
import os
import shutil
import uuid
from tool_registry import TOOLS, get_categories
from tools.pdfayir import split_pdf_by_range, split_pdf_to_single_pages
from tools.pdfbirlestir import merge_pdfs
from tools.pdfdondur import rotate_pdf
from tools.pdfekle import insert_pdf
from tools.pdfcikar import remove_pdf_pages
from tools.kdviadesi_kontrol import process_kdviadesi_pdf
from tools.kdvermedi_sifir import process_vermedi_sifir_pdf
from tools.xml_to_excel import xml_to_excel
from tools.irsaliye_no import irsaliye_no_to_excel
from tools.irsaliye_xml_to_excel import irsaliye_xml_to_excel
from tools.ekstre_boyama import paint_vakifbank_pdf
from tools.ihrac_kayitli import process_ihrac_kayitli
from tools.ithalde_indirilecek_kdv import process_ithalde_indirilecek_kdv
from tools.ithaldeindirilecekfinal import process_ithaldeindirilecekfinal
from tools.common import ensure_dirs, cleanup_old_files, secure_tr_filename, zip_files

app = Flask(__name__)
app.secret_key = "degistir-beni"

BASE_DIR = os.path.abspath(os.path.dirname(__file__))
UPLOAD_DIR = os.path.join(BASE_DIR, "uploads")
OUTPUT_DIR = os.path.join(BASE_DIR, "outputs")
RUNTIME_FILE_MAX_AGE_HOURS = 24

ensure_dirs(UPLOAD_DIR, OUTPUT_DIR)

@app.before_request
def before_request():
    cleanup_old_files(UPLOAD_DIR, max_age_hours=RUNTIME_FILE_MAX_AGE_HOURS)
    cleanup_old_files(OUTPUT_DIR, max_age_hours=RUNTIME_FILE_MAX_AGE_HOURS)


@app.route("/")
def home():
    return render_template("home.html")

@app.route("/rehber")
def rehber():
    return render_template("rehber.html")

@app.route("/ekip-araclari")
def index():
    q = request.args.get("q", "").strip().lower()
    selected_category = request.args.get("category", "").strip()
    categories = get_categories()

    filtered = TOOLS
    if selected_category:
        filtered = [t for t in filtered if t.get("category") == selected_category]

    if q:
        result = []
        for tool in filtered:
            haystack = " ".join([
                tool.get("name", ""),
                tool.get("description", ""),
                tool.get("category", ""),
                " ".join(tool.get("keywords", []))
            ]).lower()
            if q in haystack:
                result.append(tool)
        filtered = result

    return render_template(
        "index.html",
        tools=filtered,
        query=q,
        total=len(filtered),
        categories=categories,
        selected_category=selected_category
    )

@app.route("/pdf/ayir", methods=["GET", "POST"])
def pdf_ayir():
    if request.method == "POST":
        file = request.files.get("pdf_file")
        if not file or not file.filename:
            flash("Lütfen bir PDF dosyası seçin.", "error")
            return render_template("pdfayir.html")

        mode = request.form.get("mode", "range")
        safe_name = secure_tr_filename(file.filename)
        input_path = os.path.join(UPLOAD_DIR, f"{uuid.uuid4()}_{safe_name}")
        file.save(input_path)

        try:
            if mode == "range":
                page_ranges = request.form.get("page_ranges", "").strip()
                if not page_ranges:
                    raise ValueError("Sayfa aralığı girilmelidir. Örnek: 1-3,5,8-10")
                output_path = split_pdf_by_range(input_path, OUTPUT_DIR, page_ranges)
                return send_file(output_path, as_attachment=True)
            elif mode == "single":
                zip_file = split_pdf_to_single_pages(input_path, OUTPUT_DIR)
                return send_file(zip_file, as_attachment=True)
            else:
                raise ValueError("Geçersiz işlem seçildi.")
        except Exception as e:
            flash(f"Hata: {e}", "error")

    return render_template("pdfayir.html")

@app.route("/pdf/birlestir", methods=["GET", "POST"])
def pdf_birlestir():
    if request.method == "POST":
        files_in = [f for f in request.files.getlist("pdf_files") if f and f.filename]
        if len(files_in) < 2:
            flash("En az iki PDF seçmelisiniz.", "error")
            return render_template("pdfbirlestir.html")

        saved_paths = []
        for f in files_in:
            safe_name = secure_tr_filename(f.filename)
            path = os.path.join(UPLOAD_DIR, f"{uuid.uuid4()}_{safe_name}")
            f.save(path)
            saved_paths.append(path)

        try:
            output_path = merge_pdfs(saved_paths, OUTPUT_DIR)
            return send_file(output_path, as_attachment=True)
        except Exception as e:
            flash(f"Hata: {e}", "error")

    return render_template("pdfbirlestir.html")

@app.route("/pdf/dondur", methods=["GET", "POST"])
def pdf_dondur():
    if request.method == "POST":
        file = request.files.get("pdf_file")
        rotation = request.form.get("rotation", "90")
        pages = request.form.get("pages", "").strip()

        if not file or not file.filename:
            flash("Lütfen bir PDF dosyası seçin.", "error")
            return render_template("pdfdondur.html")

        safe_name = secure_tr_filename(file.filename)
        input_path = os.path.join(UPLOAD_DIR, f"{uuid.uuid4()}_{safe_name}")
        file.save(input_path)

        try:
            output_path = rotate_pdf(input_path, OUTPUT_DIR, rotation=int(rotation), pages=pages)
            return send_file(output_path, as_attachment=True)
        except Exception as e:
            flash(f"Hata: {e}", "error")

    return render_template("pdfdondur.html")

@app.route("/pdf/ekle", methods=["GET", "POST"])
def pdf_ekle():
    if request.method == "POST":
        main_file = request.files.get("main_pdf")
        insert_file = request.files.get("insert_pdf")

        if not main_file or not main_file.filename:
            flash("Lütfen ana PDF dosyasını seçin.", "error")
            return render_template("pdfekle.html")
        if not insert_file or not insert_file.filename:
            flash("Lütfen eklenecek PDF dosyasını seçin.", "error")
            return render_template("pdfekle.html")

        after_page_str = request.form.get("after_page", "0").strip()
        try:
            after_page = int(after_page_str)
        except ValueError:
            flash("Sayfa numarası geçerli bir tam sayı olmalıdır.", "error")
            return render_template("pdfekle.html")

        main_safe = secure_tr_filename(main_file.filename)
        insert_safe = secure_tr_filename(insert_file.filename)
        main_path = os.path.join(UPLOAD_DIR, f"{uuid.uuid4()}_{main_safe}")
        insert_path = os.path.join(UPLOAD_DIR, f"{uuid.uuid4()}_{insert_safe}")
        main_file.save(main_path)
        insert_file.save(insert_path)

        try:
            output_path = insert_pdf(main_path, insert_path, after_page, OUTPUT_DIR)
            return send_file(output_path, as_attachment=True)
        except Exception as e:
            flash(f"Hata: {e}", "error")

    return render_template("pdfekle.html")

@app.route("/kdv/iades-kontrol", methods=["GET", "POST"])
def kdviadesi_kontrol():
    if request.method == "POST":
        file = request.files.get("pdf_file")
        if not file or not file.filename:
            flash("Lütfen bir PDF dosyası seçin.", "error")
            return render_template("kdviadesi_kontrol.html")

        safe_name = secure_tr_filename(file.filename)
        input_path = os.path.join(UPLOAD_DIR, f"{uuid.uuid4()}_{safe_name}")
        file.save(input_path)

        try:
            status, output_path, message = process_kdviadesi_pdf(input_path, OUTPUT_DIR)
            
            if status == "error":
                flash(message, "error")
            elif status == "partial":
                flash(f"Uyarı: {message}", "warning")
                return send_file(output_path, as_attachment=True)
            else:  # success
                flash(message, "success")
                return send_file(output_path, as_attachment=True)
        except Exception as e:
            flash(f"Hata: {e}", "error")

    return render_template("kdviadesi_kontrol.html")

@app.route("/kdv/vermedi-sifir", methods=["GET", "POST"])
def kdv_vermedi_sifir():
    if request.method == "POST":
        file = request.files.get("pdf_file")
        if not file or not file.filename:
            flash("Lütfen bir PDF dosyası seçin.", "error")
            return render_template("kdvermedi_sifir.html")

        safe_name = secure_tr_filename(file.filename)
        input_path = os.path.join(UPLOAD_DIR, f"{uuid.uuid4()}_{safe_name}")
        file.save(input_path)

        try:
            status, output_path, message = process_vermedi_sifir_pdf(input_path, OUTPUT_DIR)
            
            if status == "error":
                flash(message, "error")
            elif status == "partial":
                flash(f"Uyarı: {message}", "warning")
                return send_file(output_path, as_attachment=True)
            else:  # success
                flash(message, "success")
                return send_file(output_path, as_attachment=True)
        except Exception as e:
            flash(f"Hata: {e}", "error")

    return render_template("kdvermedi_sifir.html")

@app.route("/pdf/cikar", methods=["GET", "POST"])
def pdf_cikar():
    if request.method == "POST":
        file = request.files.get("pdf_file")
        if not file or not file.filename:
            flash("Lütfen bir PDF dosyası seçin.", "error")
            return render_template("pdfcikar.html")

        pages_to_remove = request.form.get("pages_to_remove", "").strip()
        if not pages_to_remove:
            flash("Lütfen çıkarılacak sayfa numaralarını girin. Örnek: 1, 3, 5-8", "error")
            return render_template("pdfcikar.html")

        safe_name = secure_tr_filename(file.filename)
        input_path = os.path.join(UPLOAD_DIR, f"{uuid.uuid4()}_{safe_name}")
        file.save(input_path)

        try:
            output_path = remove_pdf_pages(input_path, OUTPUT_DIR, pages_to_remove)
            return send_file(output_path, as_attachment=True)
        except Exception as e:
            flash(f"Hata: {e}", "error")

    return render_template("pdfcikar.html")


@app.route("/xml/efatura", methods=["GET", "POST"])
def xml_efatura():
    if request.method == "POST":
        files = [f for f in request.files.getlist("xml_file") if f and f.filename]
        if not files:
            flash("Lütfen en az bir XML dosyası seçin.", "error")
            return render_template("xml_to_excel.html")

        saved_paths = []
        for f in files:
            safe_name = secure_tr_filename(f.filename)
            path = os.path.join(UPLOAD_DIR, f"{uuid.uuid4()}_{safe_name}")
            f.save(path)
            saved_paths.append(path)

        try:
            output_paths = [xml_to_excel(p, OUTPUT_DIR) for p in saved_paths]
            if len(output_paths) == 1:
                return send_file(output_paths[0], as_attachment=True)
            zip_path = os.path.join(OUTPUT_DIR, f"efatura_excel_{uuid.uuid4().hex[:8]}.zip")
            return send_file(zip_files(output_paths, zip_path), as_attachment=True)
        except Exception as e:
            flash(f"Hata: {e}", "error")

    return render_template("xml_to_excel.html")


@app.route("/xml/irsaliye-excel", methods=["GET", "POST"])
def xml_irsaliye_excel():
    if request.method == "POST":
        files = [f for f in request.files.getlist("xml_file") if f and f.filename]
        if not files:
            flash("Lütfen en az bir XML dosyası seçin.", "error")
            return render_template("irsaliye_xml_to_excel.html")

        saved_paths = []
        for f in files:
            safe_name = secure_tr_filename(f.filename)
            path = os.path.join(UPLOAD_DIR, f"{uuid.uuid4()}_{safe_name}")
            f.save(path)
            saved_paths.append(path)

        try:
            output_path, errors = irsaliye_xml_to_excel(saved_paths, OUTPUT_DIR)
            if errors:
                for err in errors:
                    flash(f"Uyarı: {err}", "warning")
            return send_file(output_path, as_attachment=True)
        except Exception as e:
            flash(f"Hata: {e}", "error")

    return render_template("irsaliye_xml_to_excel.html")


@app.route("/xml/irsaliye-no", methods=["GET", "POST"])
def xml_irsaliye_no():
    if request.method == "POST":
        files = [f for f in request.files.getlist("xml_file") if f and f.filename]
        if not files:
            flash("Lütfen en az bir XML dosyası seçin.", "error")
            return render_template("irsaliye_no.html")

        saved_paths = []
        for f in files:
            safe_name = secure_tr_filename(f.filename)
            path = os.path.join(UPLOAD_DIR, f"{uuid.uuid4()}_{safe_name}")
            f.save(path)
            saved_paths.append(path)

        try:
            output_path = irsaliye_no_to_excel(saved_paths, OUTPUT_DIR)
            return send_file(output_path, as_attachment=True)
        except Exception as e:
            flash(f"Hata: {e}", "error")

    return render_template("irsaliye_no.html")


@app.route("/starwood/ekstre-boyama", methods=["GET", "POST"])
def ekstre_boyama():
    if request.method == "POST":
        file = request.files.get("pdf_file")
        if not file or not file.filename:
            flash("Lütfen bir PDF dosyası seçin.", "error")
            return render_template("ekstre_boyama.html")

        original_filename = file.filename
        safe_name = secure_tr_filename(original_filename)
        input_path = os.path.join(UPLOAD_DIR, f"{uuid.uuid4()}_{safe_name}")
        file.save(input_path)

        try:
            result = paint_vakifbank_pdf(input_path, original_filename, OUTPUT_DIR)

            flash(
                f"Tamamlandı: {result['im_count']} ithalat (sarı), "
                f"{result['ex_count']} ihracat (mavi) satırı boyandı.",
                "success"
            )

            return send_file(
                result["output_path"],
                as_attachment=True,
                download_name=result["out_filename"]
            )
        except Exception as e:
            flash(f"Hata: {e}", "error")

    return render_template("ekstre_boyama.html")


@app.route("/starwood/ihrac-kayitli-hazirlama", methods=["GET", "POST"])
def ihrac_kayitli_hazirlama():
    if request.method == "POST":
        detay_file = request.files.get("detay_file")
        ozet_file = request.files.get("ozet_file")

        if not detay_file or not detay_file.filename:
            flash("Lütfen Detay Dosyası'nı seçin.", "error")
            return render_template("ihrac_kayitli.html")
        if not ozet_file or not ozet_file.filename:
            flash("Lütfen Özet Dosyası'nı seçin.", "error")
            return render_template("ihrac_kayitli.html")

        detay_name = secure_tr_filename(detay_file.filename)
        ozet_name = secure_tr_filename(ozet_file.filename)
        detay_path = os.path.join(UPLOAD_DIR, f"{uuid.uuid4()}_{detay_name}")
        ozet_path = os.path.join(UPLOAD_DIR, f"{uuid.uuid4()}_{ozet_name}")
        detay_file.save(detay_path)
        ozet_file.save(ozet_path)

        try:
            status, output_path, message = process_ihrac_kayitli(detay_path, ozet_path, OUTPUT_DIR)

            if status == "error":
                flash(message, "error")
            elif status == "partial":
                flash(f"Uyarı: {message}", "warning")
                return send_file(output_path, as_attachment=True)
            else:
                flash(message, "success")
                return send_file(output_path, as_attachment=True)
        except Exception as e:
            flash(f"Hata: {e}", "error")

    return render_template("ihrac_kayitli.html")


@app.route("/starwood/ithalde-indirilecek-kdv", methods=["GET", "POST"])
def ithalde_indirilecek_kdv():
    if request.method == "POST":
        muavin_file = request.files.get("muavin_file")
        ithalat_raporu_file = request.files.get("ithalat_raporu_file")

        if not muavin_file or not muavin_file.filename:
            flash("Lütfen Muavin dosyasını seçin.", "error")
            return render_template("ithalde_indirilecek_kdv.html")
        if not ithalat_raporu_file or not ithalat_raporu_file.filename:
            flash("Lütfen İthalat Raporu dosyasını seçin.", "error")
            return render_template("ithalde_indirilecek_kdv.html")

        muavin_name = secure_tr_filename(muavin_file.filename)
        rapor_name = secure_tr_filename(ithalat_raporu_file.filename)
        muavin_path = os.path.join(UPLOAD_DIR, f"{uuid.uuid4()}_{muavin_name}")
        rapor_path = os.path.join(UPLOAD_DIR, f"{uuid.uuid4()}_{rapor_name}")
        muavin_file.save(muavin_path)
        ithalat_raporu_file.save(rapor_path)

        try:
            status, output_path, message = process_ithalde_indirilecek_kdv(
                muavin_path,
                rapor_path,
                OUTPUT_DIR,
            )
            if status == "error":
                flash(message, "error")
            elif status == "partial":
                flash(f"Uyarı: {message}", "warning")
                return send_file(output_path, as_attachment=True)
            else:
                flash(message, "success")
                return send_file(output_path, as_attachment=True)
        except Exception as e:
            flash(f"Hata: {e}", "error")

    return render_template("ithalde_indirilecek_kdv.html")


@app.route("/starwood/ithaldeindirilecekfinal", methods=["GET", "POST"])
def ithaldeindirilecekfinal():
    if request.method == "POST":
        excel_file = request.files.get("excel_file")

        if not excel_file or not excel_file.filename:
            flash("Lütfen bir Excel dosyası seçin.", "error")
            return render_template("ithaldeindirilecekfinal.html")
        if not excel_file.filename.lower().endswith(".xlsx"):
            flash("Lütfen .xlsx uzantılı bir Excel dosyası yükleyin.", "error")
            return render_template("ithaldeindirilecekfinal.html")

        safe_name = secure_tr_filename(excel_file.filename)
        input_path = os.path.join(UPLOAD_DIR, f"{uuid.uuid4()}_{safe_name}")
        excel_file.save(input_path)

        try:
            status, output_path, message = process_ithaldeindirilecekfinal(
                input_path,
                OUTPUT_DIR,
            )
            if status == "error":
                flash(message, "error")
            else:
                flash(message, "success")
                return send_file(
                    output_path,
                    as_attachment=True,
                    download_name=os.path.basename(output_path),
                )
        except Exception as exc:
            flash(f"Hata: {exc}", "error")

    return render_template("ithaldeindirilecekfinal.html")


if __name__ == "__main__":
    app.run(debug=True)

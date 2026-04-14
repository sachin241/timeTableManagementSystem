"""
views.py — Request handlers for the AI Timetable Scheduler (v4).

v4 Endpoints (new):
  GET  /api/explain/<section>/<day>/<slot>/ → Explain why a slot was assigned
  POST /api/debug-schedule/                → Debug mode: structured failure info
  GET  /api/teacher-load/                  → Teacher hours + overload dashboard
  POST /api/timetable/partial-regen/       → Localized re-scheduling on edit

Inherited v3 Endpoints:
  GET  /                       → Landing page
  POST /upload/                → Accept CSV/XLSX/PDF/DOCX
  POST /generate/              → Generate timetables
  GET  /timetable/             → Timetable grid page
  GET  /api/timetable/         → JSON timetable (section-filtered)
  POST /api/timetable/edit/    → Edit entry with validation
  POST /api/timetable/validate/ → Dry-run validate
  GET  /download-pdf/          → PDF export
  GET  /download-xlsx/         → Excel export
  GET  /dashboard/             → Teacher load dashboard page
"""

import json
import os
from io import BytesIO

from django.conf import settings
from django.http import JsonResponse, HttpResponse, Http404
from django.shortcuts import render
from django.views.decorators.csrf import csrf_exempt
from django.views.decorators.http import require_http_methods
from django.db.models import Count, Sum, Q

from .models import Subject, Teacher, TimeSlot, Timetable, ClassSection
from .parser import parse_and_store
from .scheduler import (
    generate_time_slots,
    persist_time_slots,
    generate_timetable,
    debug_schedule,
    partial_regenerate,
    SchedulingError,
    DEFAULT_CONFIG,
)

ALLOWED_EXTENSIONS = ('.csv', '.xlsx', '.pdf', '.docx')


# ---------------------------------------------------------------------------
# Landing Page
# ---------------------------------------------------------------------------

def index(request):
    """Render the main upload page with config panel."""
    ctx = {
        'default_start': DEFAULT_CONFIG['start_time'],
        'default_end': DEFAULT_CONFIG['end_time'],
        'default_duration': DEFAULT_CONFIG['slot_duration'],
        'default_break_start': DEFAULT_CONFIG['break_start'],
        'default_break_end': DEFAULT_CONFIG['break_end'],
        'default_days': ','.join(DEFAULT_CONFIG['days']),
        'days_list': ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'],
        'default_days_set': set(DEFAULT_CONFIG['days']),
    }
    return render(request, 'scheduler/index.html', ctx)


# ---------------------------------------------------------------------------
# Upload Endpoint (v3 — multi-format)
# ---------------------------------------------------------------------------

@csrf_exempt
@require_http_methods(["POST"])
def upload_file(request):
    """
    POST /upload/
    Accepts CSV, XLSX, PDF, or DOCX files under key 'csv_file'.
    Parses and stores subjects/teachers/sections.
    """
    if 'csv_file' not in request.FILES:
        return JsonResponse({'success': False, 'error': 'No file uploaded.'}, status=400)

    uploaded = request.FILES['csv_file']
    ext = os.path.splitext(uploaded.name)[1].lower()

    if ext not in ALLOWED_EXTENSIONS:
        return JsonResponse({
            'success': False,
            'error': f'Unsupported format "{ext}". Supported: {", ".join(ALLOWED_EXTENSIONS)}',
        }, status=400)

    # Save to MEDIA_ROOT/uploads/
    upload_dir = os.path.join(settings.MEDIA_ROOT, 'uploads')
    os.makedirs(upload_dir, exist_ok=True)
    file_path = os.path.join(upload_dir, f'input{ext}')
    with open(file_path, 'wb+') as f:
        for chunk in uploaded.chunks():
            f.write(chunk)

    try:
        data = parse_and_store(file_path)
        request.session['upload_path'] = file_path
        total_hours = sum(data['subject_hours_map'].values())
        labs = [s for s in data['subjects'] if s.subject_type == 'lab']
        sections = [s.name for s in data['sections']]

        return JsonResponse({
            'success': True,
            'message': f'Parsed successfully ({ext.upper().strip(".")} format).',
            'subjects': len(data['subjects']),
            'teachers': len(data['teachers']),
            'total_hours': total_hours,
            'labs_count': len(labs),
            'sections': sections,
        })
    except ValueError as e:
        return JsonResponse({'success': False, 'error': str(e)}, status=400)
    except Exception as e:
        return JsonResponse({'success': False, 'error': f'Unexpected error: {str(e)}'}, status=500)


# ---------------------------------------------------------------------------
# Generate Endpoint (v3 — multi-section)
# ---------------------------------------------------------------------------

@csrf_exempt
@require_http_methods(["POST"])
def generate(request):
    """
    POST /generate/
    JSON body: { start_time, end_time, slot_duration, break_start, break_end, days, sections }
    """
    upload_path = request.session.get('upload_path') or request.session.get('csv_path')
    if not upload_path or not os.path.exists(upload_path):
        return JsonResponse({
            'success': False,
            'error': 'No uploaded file found. Please upload a file first.',
        }, status=400)

    # Parse config
    config = {}
    if request.content_type and 'application/json' in request.content_type:
        try:
            body = json.loads(request.body)
        except json.JSONDecodeError:
            body = {}
    else:
        body = request.POST.dict()

    for key in ('start_time', 'end_time', 'break_start', 'break_end'):
        val = body.get(key, '').strip() if isinstance(body.get(key, ''), str) else ''
        if val:
            config[key] = val

    if body.get('slot_duration'):
        try:
            config['slot_duration'] = int(body['slot_duration'])
        except (ValueError, TypeError):
            return JsonResponse({'success': False, 'error': '"slot_duration" must be integer.'}, status=400)

    days_raw = body.get('days', '')
    if days_raw:
        if isinstance(days_raw, list):
            config['days'] = [d.strip() for d in days_raw if d.strip()]
        elif isinstance(days_raw, str):
            config['days'] = [d.strip() for d in days_raw.split(',') if d.strip()]

    try:
        # Step 1: Re-parse file
        data = parse_and_store(upload_path)

        # Step 2: Generate + persist time slots
        slot_dicts = generate_time_slots(config)
        if not slot_dicts:
            return JsonResponse({
                'success': False,
                'error': 'No time slots generated. Check config.',
            }, status=400)
        orm_slots = persist_time_slots(slot_dicts)

        # Step 3: Run scheduler (auto-detects single vs multi-section)
        success = generate_timetable(data, orm_slots)

        if success:
            count = Timetable.objects.count()
            sections = list(ClassSection.objects.values_list('name', flat=True))
            return JsonResponse({
                'success': True,
                'message': f'Timetable generated: {count} sessions across {len(sections)} section(s).',
                'count': count,
                'total_slots': len(orm_slots),
                'sections': sections,
            })
        else:
            return JsonResponse({
                'success': False,
                'error': 'Scheduler could not satisfy all constraints.',
            }, status=422)

    except SchedulingError as e:
        return JsonResponse({'success': False, 'error': str(e)}, status=422)
    except ValueError as e:
        return JsonResponse({'success': False, 'error': str(e)}, status=400)
    except Exception as e:
        return JsonResponse({'success': False, 'error': f'Unexpected error: {str(e)}'}, status=500)


# ---------------------------------------------------------------------------
# Timetable View Page
# ---------------------------------------------------------------------------

def timetable_view(request):
    has_data = Timetable.objects.exists()
    sections = list(ClassSection.objects.values_list('name', flat=True))
    return render(request, 'scheduler/timetable.html', {
        'has_data': has_data,
        'sections': sections,
    })


# ---------------------------------------------------------------------------
# Timetable JSON API (v3 — section filtering)
# ---------------------------------------------------------------------------

def api_timetable(request):
    """
    GET /api/timetable/?section=CSE-A
    Returns timetable JSON for a specific section.
    If no section param, returns first available section.
    Also returns sections_list for the frontend switcher.
    """
    all_sections = list(ClassSection.objects.values_list('name', flat=True))
    requested = request.GET.get('section', '')

    entries = Timetable.objects.select_related('subject', 'teacher', 'time_slot', 'section')

    if requested and requested in all_sections:
        entries = entries.filter(section__name=requested)
    elif all_sections:
        entries = entries.filter(section__name=all_sections[0])

    if not entries.exists():
        return JsonResponse({
            'days': [], 'slots': [], 'grid': {}, 'summary': {},
            'sections_list': all_sections,
            'current_section': requested or (all_sections[0] if all_sections else ''),
        })

    days_ordered = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday']
    slot_labels_set = set()
    grid = {}
    subject_count = {}

    colors = [
        '#6366f1', '#8b5cf6', '#ec4899', '#f59e0b',
        '#10b981', '#3b82f6', '#ef4444', '#14b8a6',
        '#f97316', '#84cc16',
    ]
    subject_colors = {}
    color_idx = 0

    current_section = ''
    for entry in entries:
        slot = entry.time_slot
        label = f"{slot.start_time.strftime('%H:%M')}–{slot.end_time.strftime('%H:%M')}"
        slot_labels_set.add((slot.slot_index, label))
        current_section = entry.section.name

        if entry.subject.name not in subject_colors:
            subject_colors[entry.subject.name] = colors[color_idx % len(colors)]
            color_idx += 1

        grid.setdefault(slot.day, {})
        grid[slot.day][label] = {
            'subject': entry.subject.name,
            'teacher': entry.teacher.name,
            'color': subject_colors[entry.subject.name],
            'is_lab': entry.subject.subject_type == 'lab',
            'is_continuation': entry.is_lab_continuation,
            'entry_id': entry.id,
        }
        subject_count[entry.subject.name] = subject_count.get(entry.subject.name, 0) + 1

    sorted_labels = [l for _, l in sorted(slot_labels_set)]
    active_days = [d for d in days_ordered if d in grid]

    return JsonResponse({
        'days': active_days,
        'slots': sorted_labels,
        'grid': grid,
        'summary': {'total': entries.count(), 'subjects': subject_count},
        'sections_list': all_sections,
        'current_section': current_section,
    })


# ---------------------------------------------------------------------------
# Edit Endpoint (v3 — new)
# ---------------------------------------------------------------------------

@csrf_exempt
@require_http_methods(["POST"])
def api_edit(request):
    """
    POST /api/timetable/edit/
    JSON: { entry_id, subject_id, teacher_id }
    Changes the subject/teacher of an existing timetable entry.
    Validates constraints before saving.
    """
    try:
        body = json.loads(request.body)
    except json.JSONDecodeError:
        return JsonResponse({'success': False, 'error': 'Invalid JSON.'}, status=400)

    entry_id = body.get('entry_id')
    new_subject_id = body.get('subject_id')
    new_teacher_id = body.get('teacher_id')

    if not entry_id:
        return JsonResponse({'success': False, 'error': 'entry_id is required.'}, status=400)

    try:
        entry = Timetable.objects.select_related('time_slot', 'section').get(pk=entry_id)
    except Timetable.DoesNotExist:
        return JsonResponse({'success': False, 'error': 'Entry not found.'}, status=404)

    # Resolve new subject/teacher
    try:
        new_subject = Subject.objects.get(pk=new_subject_id) if new_subject_id else entry.subject
        new_teacher = Teacher.objects.get(pk=new_teacher_id) if new_teacher_id else entry.teacher
    except (Subject.DoesNotExist, Teacher.DoesNotExist):
        return JsonResponse({'success': False, 'error': 'Subject or teacher not found.'}, status=404)

    # ── Validate constraints ──
    conflicts = _check_conflicts(entry.time_slot, entry.section, new_teacher, exclude_id=entry.id)
    if conflicts:
        return JsonResponse({
            'success': False,
            'error': 'Conflict detected.',
            'conflicts': conflicts,
        }, status=409)

    # Apply edit
    entry.subject = new_subject
    entry.teacher = new_teacher
    entry.save()

    return JsonResponse({
        'success': True,
        'message': f'Updated: {entry.time_slot} → {new_subject.name} ({new_teacher.name})',
    })


# ---------------------------------------------------------------------------
# Validate Endpoint (v3 — dry-run)
# ---------------------------------------------------------------------------

@csrf_exempt
@require_http_methods(["POST"])
def api_validate(request):
    """
    POST /api/timetable/validate/
    Same as edit but doesn't persist — returns whether the edit would be valid.
    """
    try:
        body = json.loads(request.body)
    except json.JSONDecodeError:
        return JsonResponse({'success': False, 'error': 'Invalid JSON.'}, status=400)

    entry_id = body.get('entry_id')
    new_teacher_id = body.get('teacher_id')

    if not entry_id:
        return JsonResponse({'success': False, 'error': 'entry_id required.'}, status=400)

    try:
        entry = Timetable.objects.select_related('time_slot', 'section').get(pk=entry_id)
    except Timetable.DoesNotExist:
        return JsonResponse({'success': False, 'error': 'Entry not found.'}, status=404)

    new_teacher = entry.teacher
    if new_teacher_id:
        try:
            new_teacher = Teacher.objects.get(pk=new_teacher_id)
        except Teacher.DoesNotExist:
            return JsonResponse({'success': False, 'error': 'Teacher not found.'}, status=404)

    conflicts = _check_conflicts(entry.time_slot, entry.section, new_teacher, exclude_id=entry.id)

    return JsonResponse({
        'valid': len(conflicts) == 0,
        'conflicts': conflicts,
    })


def _check_conflicts(time_slot, section, teacher, exclude_id=None):
    """
    Check for teacher conflicts across ALL sections for a given time_slot.
    Returns list of conflict strings.
    """
    conflicts = []

    # Teacher busy in another section at same time?
    qs = Timetable.objects.filter(
        time_slot=time_slot,
        teacher=teacher,
    ).exclude(pk=exclude_id or -1)

    for conflict_entry in qs:
        conflicts.append(
            f"{teacher.name} is already teaching {conflict_entry.subject.name} "
            f"in [{conflict_entry.section.name}] at {time_slot}"
        )

    return conflicts


# ---------------------------------------------------------------------------
# Explain Endpoint (v4)
# ---------------------------------------------------------------------------

def api_explain(request, section, day, slot):
    """
    GET /api/explain/<section>/<day>/<slot>/
    Returns the explanation metadata for a specific timetable cell.
    """
    # Parse slot label (e.g. "09:00–10:00")
    slot_label = slot.replace('-', '–')  # normalise dash vs en-dash

    try:
        entry = Timetable.objects.select_related(
            'subject', 'teacher', 'time_slot', 'section'
        ).filter(
            section__name=section,
            time_slot__day=day,
        ).first()

        # Find by matching slot label
        entries = Timetable.objects.select_related(
            'subject', 'teacher', 'time_slot', 'section'
        ).filter(
            section__name=section,
            time_slot__day=day,
        )

        for e in entries:
            label = f"{e.time_slot.start_time.strftime('%H:%M')}–{e.time_slot.end_time.strftime('%H:%M')}"
            if label == slot_label or label.replace('–', '-') == slot.replace('–', '-'):
                return JsonResponse({
                    'found': True,
                    'subject': e.subject.name,
                    'teacher': e.teacher.name,
                    'slot': label,
                    'day': day,
                    'section': section,
                    'explanation': e.explanation or {},
                    'is_lab': e.subject.subject_type == 'lab',
                    'is_continuation': e.is_lab_continuation,
                })

    except Exception:
        pass

    return JsonResponse({'found': False, 'explanation': {}})


# ---------------------------------------------------------------------------
# Debug Schedule Endpoint (v4)
# ---------------------------------------------------------------------------

@csrf_exempt
@require_http_methods(["POST"])
def api_debug_schedule(request):
    """
    POST /api/debug-schedule/
    Re-runs scheduler in debug mode without persisting.
    Returns structured diagnostics.
    """
    upload_path = request.session.get('upload_path') or request.session.get('csv_path')
    if not upload_path or not os.path.exists(upload_path):
        return JsonResponse({
            'success': False, 'error': 'No uploaded file. Upload first.'
        }, status=400)

    # Parse config from body
    config = {}
    if request.content_type and 'application/json' in request.content_type:
        try:
            body = json.loads(request.body)
        except json.JSONDecodeError:
            body = {}
    else:
        body = {}

    for key in ('start_time', 'end_time', 'break_start', 'break_end'):
        val = body.get(key, '')
        if val:
            config[key] = val
    if body.get('slot_duration'):
        config['slot_duration'] = int(body['slot_duration'])
    days_raw = body.get('days', '')
    if days_raw:
        config['days'] = days_raw if isinstance(days_raw, list) else [d.strip() for d in days_raw.split(',')]

    try:
        data = parse_and_store(upload_path)
        slot_dicts = generate_time_slots(config)
        orm_slots = persist_time_slots(slot_dicts)
        reports = debug_schedule(data, orm_slots)

        return JsonResponse({
            'success': True,
            'reports': reports,
        })
    except Exception as e:
        return JsonResponse({'success': False, 'error': str(e)}, status=500)


# ---------------------------------------------------------------------------
# Teacher Load Dashboard (v4)
# ---------------------------------------------------------------------------

def api_teacher_load(request):
    """
    GET /api/teacher-load/
    Returns teacher workload data: assigned hours, utilization, overload status.
    """
    total_slots = TimeSlot.objects.count()
    n_sections = ClassSection.objects.count() or 1

    teachers = Teacher.objects.all()
    result = []

    for teacher in teachers:
        entries = Timetable.objects.filter(teacher=teacher)
        assigned_hours = entries.count()

        # Per-section breakdown
        section_breakdown = {}
        for entry in entries.select_related('section'):
            sec = entry.section.name
            section_breakdown[sec] = section_breakdown.get(sec, 0) + 1

        # Max available = total slots (teacher could teach every slot)
        max_available = total_slots
        utilization = round((assigned_hours / max_available * 100), 1) if max_available else 0

        # Overload threshold: >80% utilization
        overload = utilization > 80

        result.append({
            'id': teacher.id,
            'name': teacher.name,
            'assigned_hours': assigned_hours,
            'max_available': max_available,
            'utilization': utilization,
            'overloaded': overload,
            'sections': section_breakdown,
            'subjects': list(teacher.subjects.values_list('name', flat=True)),
        })

    # Sort: overloaded first, then by utilization desc
    result.sort(key=lambda x: (-x['overloaded'], -x['utilization']))

    return JsonResponse({
        'teachers': result,
        'total_slots': total_slots,
        'total_sections': n_sections,
    })


def dashboard_view(request):
    """Render the teacher load dashboard page."""
    return render(request, 'scheduler/dashboard.html')


# ---------------------------------------------------------------------------
# Partial Regeneration Endpoint (v4)
# ---------------------------------------------------------------------------

@csrf_exempt
@require_http_methods(["POST"])
def api_partial_regen(request):
    """
    POST /api/timetable/partial-regen/
    JSON: { entry_id, subject_id, teacher_id }
    Localized re-scheduling: apply edit + resolve cascading conflicts.
    """
    try:
        body = json.loads(request.body)
    except json.JSONDecodeError:
        return JsonResponse({'success': False, 'error': 'Invalid JSON.'}, status=400)

    entry_id = body.get('entry_id')
    if not entry_id:
        return JsonResponse({'success': False, 'error': 'entry_id required.'}, status=400)

    try:
        result = partial_regenerate(
            entry_id=int(entry_id),
            new_subject_id=body.get('subject_id'),
            new_teacher_id=body.get('teacher_id'),
        )
        status_code = 200 if result['success'] else 422
        return JsonResponse(result, status=status_code)
    except Exception as e:
        return JsonResponse({'success': False, 'error': str(e)}, status=500)


# ---------------------------------------------------------------------------
# PDF Download (v3 — section filtering)
# ---------------------------------------------------------------------------

def download_pdf(request):
    """GET /download-pdf/?section=CSE-A — PDF export."""
    from reportlab.lib.pagesizes import landscape, A3
    from reportlab.lib import colors as rl_colors
    from reportlab.lib.units import cm
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.enums import TA_CENTER

    section_name = request.GET.get('section', '')
    entries = Timetable.objects.select_related('subject', 'teacher', 'time_slot', 'section')

    if section_name:
        entries = entries.filter(section__name=section_name)

    if not entries.exists():
        raise Http404("No timetable data found.")

    # Group by section
    sections_data = {}
    for entry in entries:
        sec = entry.section.name
        sections_data.setdefault(sec, []).append(entry)

    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=landscape(A3), topMargin=1.5*cm, bottomMargin=1.5*cm)
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle('title', parent=styles['Heading1'], alignment=TA_CENTER, fontSize=20)
    cell_style = ParagraphStyle('cell', parent=styles['Normal'], fontSize=8, alignment=TA_CENTER)

    story = []
    days_ordered = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday']

    for sec_name, sec_entries in sorted(sections_data.items()):
        slot_set = {}
        grid = {}
        for e in sec_entries:
            slot = e.time_slot
            label = f"{slot.start_time.strftime('%H:%M')}–{slot.end_time.strftime('%H:%M')}"
            slot_set[slot.slot_index] = label
            text = f"{e.subject.name}\n({e.teacher.name})"
            if e.subject.subject_type == 'lab':
                text = f"🔬 {text}"
            grid.setdefault(slot.day, {})[label] = text

        sorted_labels = [v for _, v in sorted(slot_set.items())]
        active_days = [d for d in days_ordered if d in grid]

        header = ['Time Slot'] + active_days
        table_data = [header]
        for sl in sorted_labels:
            row = [sl]
            for day in active_days:
                cell_text = grid.get(day, {}).get(sl, '—')
                row.append(Paragraph(cell_text.replace('\n', '<br/>'), cell_style))
            table_data.append(row)

        col_widths = [3 * cm] + [3.5 * cm] * len(active_days)
        table = Table(table_data, colWidths=col_widths, repeatRows=1)
        table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), rl_colors.HexColor('#4f46e5')),
            ('TEXTCOLOR', (0, 0), (-1, 0), rl_colors.white),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 10),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('GRID', (0, 0), (-1, -1), 0.5, rl_colors.grey),
            ('BACKGROUND', (0, 1), (0, -1), rl_colors.HexColor('#e0e7ff')),
            ('ROWBACKGROUNDS', (1, 1), (-1, -1), [rl_colors.white, rl_colors.HexColor('#f5f3ff')]),
            ('TOPPADDING', (0, 0), (-1, -1), 6),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
        ]))

        story.append(Paragraph(f'Timetable — {sec_name}', title_style))
        story.append(Spacer(1, 0.5 * cm))
        story.append(table)
        story.append(Spacer(1, 1.5 * cm))

    doc.build(story)
    buffer.seek(0)
    response = HttpResponse(buffer, content_type='application/pdf')
    response['Content-Disposition'] = 'attachment; filename="timetable.pdf"'
    return response


# ---------------------------------------------------------------------------
# Excel Download (v3 — new)
# ---------------------------------------------------------------------------

def download_xlsx(request):
    """GET /download-xlsx/?section=all — Excel export."""
    from openpyxl import Workbook
    from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

    section_name = request.GET.get('section', 'all')
    entries = Timetable.objects.select_related('subject', 'teacher', 'time_slot', 'section')

    if section_name and section_name != 'all':
        entries = entries.filter(section__name=section_name)

    if not entries.exists():
        raise Http404("No timetable data found.")

    # Group by section
    sections_data = {}
    for entry in entries:
        sec = entry.section.name
        sections_data.setdefault(sec, []).append(entry)

    wb = Workbook()
    wb.remove(wb.active)  # Remove default sheet

    days_ordered = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday']
    header_fill = PatternFill(start_color='4F46E5', end_color='4F46E5', fill_type='solid')
    header_font = Font(bold=True, color='FFFFFF', size=11)
    cell_align = Alignment(horizontal='center', vertical='center', wrap_text=True)
    thin_border = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'), bottom=Side(style='thin'),
    )

    for sec_name in sorted(sections_data.keys()):
        sec_entries = sections_data[sec_name]
        ws = wb.create_sheet(title=sec_name[:31])  # Excel 31-char limit

        slot_set = {}
        grid = {}
        for e in sec_entries:
            slot = e.time_slot
            label = f"{slot.start_time.strftime('%H:%M')}–{slot.end_time.strftime('%H:%M')}"
            slot_set[slot.slot_index] = label
            tag = ' [LAB]' if e.subject.subject_type == 'lab' else ''
            grid.setdefault(slot.day, {})[label] = f"{e.subject.name}{tag}\n{e.teacher.name}"

        sorted_labels = [v for _, v in sorted(slot_set.items())]
        active_days = [d for d in days_ordered if d in grid]

        # Header row
        headers = ['Time Slot'] + active_days
        for col_idx, hdr in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col_idx, value=hdr)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = cell_align
            cell.border = thin_border

        # Data rows
        for row_idx, sl in enumerate(sorted_labels, 2):
            ws.cell(row=row_idx, column=1, value=sl).alignment = cell_align
            ws.cell(row=row_idx, column=1).border = thin_border
            ws.cell(row=row_idx, column=1).font = Font(bold=True)
            for col_idx, day in enumerate(active_days, 2):
                val = grid.get(day, {}).get(sl, '—')
                cell = ws.cell(row=row_idx, column=col_idx, value=val)
                cell.alignment = cell_align
                cell.border = thin_border

        # Column widths
        ws.column_dimensions['A'].width = 16
        for i in range(2, len(active_days) + 2):
            ws.column_dimensions[chr(64 + i)].width = 22

    buffer = BytesIO()
    wb.save(buffer)
    buffer.seek(0)

    response = HttpResponse(
        buffer,
        content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    )
    response['Content-Disposition'] = 'attachment; filename="timetable.xlsx"'
    return response

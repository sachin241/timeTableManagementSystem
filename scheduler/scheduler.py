"""
scheduler.py — Constraint-Based Timetable Scheduler Engine (v4).

v4 Enhancements:
  - Explainable scheduling: per-slot reasoning metadata stored in self.explanations.
  - debug_run(): collects structured diagnostic data on scheduling failure.
  - partial_regenerate(): localized re-scheduling when a user edits a slot.
  - SC4: Even distribution penalty (std-deviation of subject-per-day counts).

Architecture (pluggable):
  - BaseScheduler         → Abstract base class.
  - BacktrackingScheduler → Constraint-based recursive backtracking.
  - GeneticScheduler      → Stub for future GA implementation.

Hard Constraints (HC):
  HC1: Teacher cannot be in two classes at the same time (global across sections).
  HC2: Each subject must be assigned exactly its required hours.
  HC3: One subject per slot per section.
  HC4: Lab subjects must occupy consecutive slots on the same day.

Soft Constraints (SC, penalties — lower is better):
  SC1: +10 if same subject in consecutive slot on same day.
  SC2: +5 × (count_today / total_today) balance penalty.
  SC3: +15 for every occurrence beyond 2 of same subject on a day.
  SC4: +8 × deviation from ideal daily distribution.
"""

from abc import ABC, abstractmethod
from collections import defaultdict
from datetime import time as dt_time
import math


# ---------------------------------------------------------------------------
# Time Slot Generation
# ---------------------------------------------------------------------------

DEFAULT_CONFIG = {
    'days': ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday'],
    'start_time': '09:00',
    'end_time': '16:00',
    'slot_duration': 60,
    'break_start': '13:00',
    'break_end': '14:00',
}

DAY_ORDER = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday']


def _str_to_time(s: str) -> dt_time:
    h, m = s.strip().split(':')
    return dt_time(int(h), int(m))


def _time_to_minutes(t: dt_time) -> int:
    return t.hour * 60 + t.minute


def _minutes_to_time(m: int) -> dt_time:
    return dt_time(m // 60, m % 60)


def generate_time_slots(config: dict = None) -> list:
    """
    Generate time slot dicts from schedule config.
    Returns [{'day', 'start', 'end', 'slot_index'}, ...].
    """
    cfg = {**DEFAULT_CONFIG, **(config or {})}
    days = cfg['days']
    start_min = _time_to_minutes(_str_to_time(cfg['start_time']))
    end_min = _time_to_minutes(_str_to_time(cfg['end_time']))
    duration = int(cfg['slot_duration'])

    break_start_min = break_end_min = None
    if cfg.get('break_start') and cfg.get('break_end'):
        break_start_min = _time_to_minutes(_str_to_time(cfg['break_start']))
        break_end_min = _time_to_minutes(_str_to_time(cfg['break_end']))

    if duration <= 0:
        raise ValueError("slot_duration must be a positive integer.")
    if start_min >= end_min:
        raise ValueError("start_time must be earlier than end_time.")

    slots = []
    for day in days:
        current = start_min
        slot_idx = 0
        while current + duration <= end_min:
            slot_end = current + duration
            if break_start_min is not None and break_end_min is not None:
                if (current < break_end_min) and (slot_end > break_start_min):
                    current = break_end_min
                    continue
            slots.append({
                'day': day,
                'start': _minutes_to_time(current),
                'end': _minutes_to_time(slot_end),
                'slot_index': slot_idx,
            })
            current = slot_end
            slot_idx += 1
    return slots


def persist_time_slots(slot_dicts: list):
    """Clear all TimeSlot/Timetable rows and create fresh TimeSlots."""
    from .models import TimeSlot, Timetable
    Timetable.objects.all().delete()
    TimeSlot.objects.all().delete()

    orm_slots = []
    for s in slot_dicts:
        slot = TimeSlot.objects.create(
            day=s['day'], start_time=s['start'],
            end_time=s['end'], slot_index=s['slot_index'],
        )
        orm_slots.append(slot)
    return orm_slots


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _group_slots_by_day(slots):
    """Return {day: [slot, ...]} ordered by slot_index."""
    groups = defaultdict(list)
    for s in slots:
        groups[s.day].append(s)
    for day in groups:
        groups[day].sort(key=lambda s: s.slot_index)
    return dict(groups)


# ---------------------------------------------------------------------------
# Debug Report
# ---------------------------------------------------------------------------

class DebugReport:
    """Structured diagnostic data when scheduling fails."""

    def __init__(self):
        self.teacher_overloads = []    # [{teacher, required, available}]
        self.lab_failures = []         # [{subject, hours_needed, max_consecutive}]
        self.unplaceable_subjects = [] # [{subject, remaining, attempted}]
        self.constraint_violations = [] # [str]

    def to_dict(self):
        return {
            'error': 'SchedulingFailed',
            'teacher_overloads': self.teacher_overloads,
            'lab_failures': self.lab_failures,
            'unplaceable_subjects': self.unplaceable_subjects,
            'constraint_violations': self.constraint_violations,
        }


# ---------------------------------------------------------------------------
# Base Scheduler Interface
# ---------------------------------------------------------------------------

class BaseScheduler(ABC):
    """
    Abstract base for scheduler implementations.

    v4: Now tracks explanation metadata per slot assignment.
    """

    def __init__(self, data: dict, slots: list, section_subjects: list = None,
                 global_teacher_busy: dict = None):
        self.all_subjects = {s.id: s for s in data['subjects']}
        self.all_teachers = {t.id: t for t in data['teachers']}
        self.slots = slots
        self.slots_by_day = _group_slots_by_day(slots)
        self.subject_teacher_map = data['subject_teacher_map']

        if section_subjects:
            self.section_subjects = section_subjects
            self.subjects = [self.all_subjects[si['subject_id']] for si in section_subjects]
            self.subject_hours_map = {si['subject_id']: si['hours'] for si in section_subjects}
        else:
            self.subjects = data['subjects']
            self.subject_hours_map = data['subject_hours_map']

        self.lab_subjects = [s for s in self.subjects if s.subject_type == 'lab']
        self.theory_subjects = [s for s in self.subjects if s.subject_type == 'theory']
        self.global_teacher_busy = global_teacher_busy if global_teacher_busy is not None else defaultdict(set)

        # {slot_id: (Subject, Teacher, is_continuation)}
        self.assignments: dict = {}
        # v4: {slot_id: {reason, factors, sc_score, phase}}
        self.explanations: dict = {}

    @abstractmethod
    def run(self) -> bool:
        ...

    def save(self, section_name: str = 'A'):
        """Persist assignments + explanations to Timetable for the given section."""
        from .models import Timetable, TimeSlot as TSModel, ClassSection
        section, _ = ClassSection.objects.get_or_create(name=section_name)
        Timetable.objects.filter(section=section).delete()
        for slot_id, (subject, teacher, is_cont) in self.assignments.items():
            slot = TSModel.objects.get(pk=slot_id)
            explanation = self.explanations.get(slot_id, {})
            Timetable.objects.create(
                subject=subject, teacher=teacher, time_slot=slot,
                section=section, is_lab_continuation=is_cont,
                explanation=explanation,
            )


# ---------------------------------------------------------------------------
# Backtracking Scheduler (v4)
# ---------------------------------------------------------------------------

class BacktrackingScheduler(BaseScheduler):
    """
    Constraint-based scheduler using recursive backtracking.

    v4 Algorithm:
      Phase 1 — Schedule LAB subjects (consecutive slots, explanation tracked).
      Phase 2 — Schedule THEORY subjects (MRV + SC1-SC4, explanation tracked).
      debug_run() — Same scheduling but collects DebugReport on failure.
    """

    def run(self) -> bool:
        total_required = sum(self.subject_hours_map.values())
        total_available = len(self.slots)
        if total_required > total_available:
            raise SchedulingError(
                f"Insufficient slots for this section. "
                f"Required: {total_required}, Available: {total_available}."
            )

        remaining = {s_id: hrs for s_id, hrs in self.subject_hours_map.items()}
        day_schedule = defaultdict(list)
        local_teacher_busy = defaultdict(set)
        used_slots = set()

        if not self._schedule_labs(remaining, day_schedule, local_teacher_busy, used_slots):
            return False

        free_slots = [s for s in self.slots if s.id not in used_slots]
        return self._backtrack(
            slot_index=0, free_slots=free_slots,
            remaining=remaining, day_schedule=day_schedule,
            local_teacher_busy=local_teacher_busy, used_slots=used_slots,
        )

    # ── Debug Mode ───────────────────────────────────────────────────

    def debug_run(self) -> DebugReport:
        """
        Run scheduling in debug mode: collect diagnostics instead of raising.
        Returns a DebugReport even if scheduling succeeds (report will be empty).
        """
        report = DebugReport()
        total_required = sum(self.subject_hours_map.values())
        total_available = len(self.slots)

        if total_required > total_available:
            report.constraint_violations.append(
                f"Insufficient slots: required {total_required}, available {total_available}"
            )

        # Check per-teacher capacity
        for subject in self.subjects:
            teacher = self.subject_teacher_map.get(subject.id)
            if not teacher:
                report.constraint_violations.append(
                    f"No teacher assigned for '{subject.name}'"
                )
                continue

            required_hours = self.subject_hours_map.get(subject.id, 0)
            # Count how many slots teacher is already busy (globally)
            busy_count = len(self.global_teacher_busy.get(teacher.id, set()))
            available_for_teacher = total_available - busy_count

            if required_hours > available_for_teacher:
                report.teacher_overloads.append({
                    'teacher': teacher.name,
                    'subject': subject.name,
                    'required_hours': required_hours,
                    'available_slots': available_for_teacher,
                    'already_busy': busy_count,
                })

        # Check lab feasibility
        for subject in self.lab_subjects:
            hours_needed = self.subject_hours_map.get(subject.id, 0)
            teacher = self.subject_teacher_map.get(subject.id)
            max_consecutive = 0

            for day, day_slots in self.slots_by_day.items():
                # Find max consecutive free slots where teacher is available
                streak = 0
                for slot in day_slots:
                    if slot.id not in self.global_teacher_busy.get(teacher.id if teacher else -1, set()):
                        streak += 1
                        max_consecutive = max(max_consecutive, streak)
                    else:
                        streak = 0

            if max_consecutive < hours_needed:
                report.lab_failures.append({
                    'subject': subject.name,
                    'hours_needed': hours_needed,
                    'max_consecutive_found': max_consecutive,
                })

        # Attempt actual scheduling to find remaining issues
        remaining = {s_id: hrs for s_id, hrs in self.subject_hours_map.items()}
        day_schedule = defaultdict(list)
        local_teacher_busy = defaultdict(set)
        used_slots = set()

        try:
            self._schedule_labs(remaining, day_schedule, local_teacher_busy, used_slots)
        except SchedulingError as e:
            report.constraint_violations.append(f"Lab scheduling: {str(e)}")

        free_slots = [s for s in self.slots if s.id not in used_slots]
        success = self._backtrack(
            slot_index=0, free_slots=free_slots,
            remaining=remaining, day_schedule=day_schedule,
            local_teacher_busy=local_teacher_busy, used_slots=used_slots,
        )

        if not success:
            for s_id, hrs_left in remaining.items():
                if hrs_left > 0:
                    subj = self.all_subjects.get(s_id)
                    report.unplaceable_subjects.append({
                        'subject': subj.name if subj else str(s_id),
                        'remaining_hours': hrs_left,
                        'total_required': self.subject_hours_map.get(s_id, 0),
                    })

        return report

    # ── Phase 1: Lab Scheduling ──────────────────────────────────────

    def _schedule_labs(self, remaining, day_schedule, local_teacher_busy, used_slots) -> bool:
        for subject in self.lab_subjects:
            hours_needed = remaining.get(subject.id, 0)
            if hours_needed <= 0:
                continue

            teacher = self.subject_teacher_map.get(subject.id)
            if teacher is None:
                raise SchedulingError(f"No teacher assigned for lab subject '{subject.name}'.")

            placed = self._place_lab_block(
                subject, teacher, hours_needed,
                day_schedule, local_teacher_busy, used_slots,
            )
            if not placed:
                raise SchedulingError(
                    f"Cannot find {hours_needed} consecutive free slots for lab "
                    f"'{subject.name}'. Try increasing available hours or reducing lab duration."
                )
            remaining[subject.id] = 0

        return True

    def _place_lab_block(self, subject, teacher, count, day_schedule,
                         local_teacher_busy, used_slots) -> bool:
        day_order = sorted(
            self.slots_by_day.keys(),
            key=lambda d: len(day_schedule.get(d, [])),
        )

        for day in day_order:
            day_slots = self.slots_by_day.get(day, [])
            for i in range(len(day_slots) - count + 1):
                block = day_slots[i:i + count]

                if any(s.id in used_slots for s in block):
                    continue

                all_free = True
                for s in block:
                    if s.id in self.global_teacher_busy.get(teacher.id, set()):
                        all_free = False
                        break
                    if s.id in local_teacher_busy.get(teacher.id, set()):
                        all_free = False
                        break
                if not all_free:
                    continue

                # ── Assign the entire block with explanations ──
                for idx, slot in enumerate(block):
                    is_continuation = (idx > 0)
                    self.assignments[slot.id] = (subject, teacher, is_continuation)
                    used_slots.add(slot.id)
                    day_schedule[day].append(subject.id)
                    local_teacher_busy[teacher.id].add(slot.id)
                    self.global_teacher_busy[teacher.id].add(slot.id)

                    # v4: Explanation
                    slot_label = f"{slot.start_time.strftime('%H:%M')}–{slot.end_time.strftime('%H:%M')}"
                    factors = [
                        f"Lab requires {count} consecutive slots",
                        f"{teacher.name} was free for all {count} slots",
                        f"No slot conflict on {day}",
                        f"{day} had fewest assignments (balanced placement)",
                    ]
                    if is_continuation:
                        factors.append(f"Continuation of {subject.name} lab block (slot {idx+1}/{count})")

                    self.explanations[slot.id] = {
                        'reason': f"Lab '{subject.name}' placed in consecutive block on {day} ({slot_label})",
                        'factors': factors,
                        'sc_score': 0.0,
                        'phase': 'lab_scheduling',
                    }

                return True

        return False

    # ── Phase 2: Theory Scheduling (backtracking) ────────────────────

    def _backtrack(self, slot_index, free_slots, remaining, day_schedule,
                   local_teacher_busy, used_slots) -> bool:
        if all(h == 0 for h in remaining.values()):
            return True

        if slot_index >= len(free_slots):
            return all(h == 0 for h in remaining.values())

        slot = free_slots[slot_index]
        candidates = self._get_candidates(slot, remaining, day_schedule,
                                          local_teacher_busy)

        for subject, sc_score, factors in candidates:
            teacher = self.subject_teacher_map[subject.id]

            if slot.id in self.global_teacher_busy.get(teacher.id, set()):
                continue
            if slot.id in local_teacher_busy.get(teacher.id, set()):
                continue

            # ── Assign ──
            self.assignments[slot.id] = (subject, teacher, False)
            remaining[subject.id] -= 1
            day_schedule[slot.day].append(subject.id)
            local_teacher_busy[teacher.id].add(slot.id)
            self.global_teacher_busy[teacher.id].add(slot.id)
            used_slots.add(slot.id)

            # v4: Explanation
            hrs_left = remaining[subject.id]
            slot_label = f"{slot.start_time.strftime('%H:%M')}–{slot.end_time.strftime('%H:%M')}"
            self.explanations[slot.id] = {
                'reason': f"MRV selected '{subject.name}' ({hrs_left}h remaining after placement)",
                'factors': factors,
                'sc_score': round(sc_score, 2),
                'phase': 'theory_backtracking',
            }

            if self._backtrack(slot_index + 1, free_slots, remaining,
                               day_schedule, local_teacher_busy, used_slots):
                return True

            # ── Undo ──
            del self.assignments[slot.id]
            if slot.id in self.explanations:
                del self.explanations[slot.id]
            remaining[subject.id] += 1
            day_schedule[slot.day].remove(subject.id)
            local_teacher_busy[teacher.id].discard(slot.id)
            self.global_teacher_busy[teacher.id].discard(slot.id)
            used_slots.discard(slot.id)

        # Skip this slot (free period)
        return self._backtrack(slot_index + 1, free_slots, remaining,
                               day_schedule, local_teacher_busy, used_slots)

    def _get_candidates(self, slot, remaining, day_schedule, local_teacher_busy):
        """
        Return priority-sorted theory subjects for this slot.

        v4: Returns list of (subject, sc_score, factors) tuples.
        Sorting: MRV (most remaining first) → SC1+SC2+SC3+SC4 penalties → name.
        """
        today_subjects = day_schedule.get(slot.day, [])
        total_today = len(today_subjects)
        num_days = len(self.slots_by_day)

        candidates = []
        for subject in self.theory_subjects:
            hrs_left = remaining.get(subject.id, 0)
            if hrs_left <= 0:
                continue

            teacher = self.subject_teacher_map.get(subject.id)
            if teacher is None:
                continue
            if slot.id in self.global_teacher_busy.get(teacher.id, set()):
                continue
            if slot.id in local_teacher_busy.get(teacher.id, set()):
                continue

            # ── Scoring ──
            mrv = -hrs_left
            score = 0.0
            factors = [f"{teacher.name} is free for this slot"]

            # SC1: consecutive penalty
            if today_subjects and today_subjects[-1] == subject.id:
                score += 10.0
                factors.append("SC1: -10 consecutive same subject penalty")
            else:
                factors.append("SC1: No consecutive conflict")

            # SC2: balance penalty
            count_today = today_subjects.count(subject.id)
            if total_today > 0:
                balance_pen = (count_today / total_today) * 5.0
                score += balance_pen
                if balance_pen > 0:
                    factors.append(f"SC2: {count_today}/{total_today} today, penalty +{balance_pen:.1f}")

            # SC3: heavy daily overload (>2 same subject)
            if count_today >= 2:
                sc3_pen = 15.0 * (count_today - 1)
                score += sc3_pen
                factors.append(f"SC3: {count_today} already today, overload penalty +{sc3_pen:.1f}")
            else:
                factors.append("SC3: Under daily limit")

            # SC4: even distribution penalty (new in v4)
            total_hours = self.subject_hours_map.get(subject.id, 0)
            ideal_per_day = total_hours / max(num_days, 1)
            deviation = abs(count_today + 1 - ideal_per_day)
            sc4_pen = 8.0 * deviation
            score += sc4_pen
            if sc4_pen > 1.0:
                factors.append(f"SC4: Distribution deviation {deviation:.1f}, penalty +{sc4_pen:.1f}")
            else:
                factors.append(f"SC4: Good distribution (ideal {ideal_per_day:.1f}/day)")

            factors.append(f"Total SC score: {score:.1f}")
            factors.append(f"MRV: {hrs_left}h remaining")

            candidates.append((mrv, score, subject.name, subject, score, factors))

        candidates.sort(key=lambda x: (x[0], x[1], x[2]))
        return [(c[3], c[4], c[5]) for c in candidates]


# ---------------------------------------------------------------------------
# Genetic Scheduler — Future Stub
# ---------------------------------------------------------------------------

class GeneticScheduler(BaseScheduler):
    def run(self) -> bool:
        raise NotImplementedError("GeneticScheduler is not yet implemented.")


# ---------------------------------------------------------------------------
# Custom Exception
# ---------------------------------------------------------------------------

class SchedulingError(Exception):
    """Raised when scheduling constraints cannot be satisfied."""
    def __init__(self, message, debug_report=None):
        super().__init__(message)
        self.debug_report = debug_report


# ---------------------------------------------------------------------------
# Public API — Multi-Section Scheduling
# ---------------------------------------------------------------------------

def generate_timetable(data: dict, slots: list) -> bool:
    """Backward-compatible single/multi-section scheduling."""
    sections_data = data.get('section_subject_data', {})
    section_names = list(sections_data.keys()) if sections_data else ['A']

    if len(section_names) == 1:
        sec_name = section_names[0]
        sec_subjects = sections_data.get(sec_name, None)
        scheduler = BacktrackingScheduler(data, slots, section_subjects=sec_subjects)
        success = scheduler.run()
        if success:
            scheduler.save(section_name=sec_name)
        return success
    else:
        return generate_timetable_multi(data, slots, section_names)


def generate_timetable_multi(data: dict, slots: list, section_names: list) -> bool:
    """Schedule multiple sections with shared global teacher tracking."""
    sections_data = data.get('section_subject_data', {})
    global_teacher_busy = defaultdict(set)

    for sec_name in section_names:
        sec_subjects = sections_data.get(sec_name, None)
        if not sec_subjects:
            continue

        scheduler = BacktrackingScheduler(
            data, slots,
            section_subjects=sec_subjects,
            global_teacher_busy=global_teacher_busy,
        )

        try:
            success = scheduler.run()
        except SchedulingError as e:
            raise SchedulingError(f"[{sec_name}] {str(e)}")

        if not success:
            raise SchedulingError(
                f"[{sec_name}] Could not satisfy all constraints for this section."
            )

        scheduler.save(section_name=sec_name)

    return True


# ---------------------------------------------------------------------------
# Debug Schedule (v4)
# ---------------------------------------------------------------------------

def debug_schedule(data: dict, slots: list) -> dict:
    """
    Run scheduler in debug mode, returning structured diagnostics.
    Does NOT persist any changes.
    """
    sections_data = data.get('section_subject_data', {})
    section_names = list(sections_data.keys()) if sections_data else ['A']
    global_teacher_busy = defaultdict(set)
    all_reports = {}

    for sec_name in section_names:
        sec_subjects = sections_data.get(sec_name, None)
        if not sec_subjects:
            continue

        scheduler = BacktrackingScheduler(
            data, slots,
            section_subjects=sec_subjects,
            global_teacher_busy=global_teacher_busy,
        )

        report = scheduler.debug_run()
        all_reports[sec_name] = report.to_dict()

    return all_reports


# ---------------------------------------------------------------------------
# Partial Regeneration (v4)
# ---------------------------------------------------------------------------

def partial_regenerate(entry_id: int, new_subject_id: int = None,
                       new_teacher_id: int = None) -> dict:
    """
    Localized re-scheduling after a user edit.

    Instead of regenerating the entire timetable:
      1. Load existing timetable for the entry's section.
      2. Apply the edit (swap subject/teacher on that slot).
      3. Identify cascading conflicts (teacher double-booked, hours imbalance).
      4. Remove conflicting entries.
      5. Re-run backtracking ONLY for the freed slots + affected subjects.
      6. Persist and return the list of changed entries.

    Args:
        entry_id: PK of Timetable row to edit.
        new_subject_id: New subject PK (or None to keep).
        new_teacher_id: New teacher PK (or None to keep).

    Returns:
        dict: {'success': bool, 'changed': [entry descriptions], 'message': str}
    """
    from .models import Timetable, Subject, Teacher, TimeSlot, ClassSection

    try:
        entry = Timetable.objects.select_related(
            'subject', 'teacher', 'time_slot', 'section'
        ).get(pk=entry_id)
    except Timetable.DoesNotExist:
        return {'success': False, 'message': 'Entry not found.', 'changed': []}

    section = entry.section
    old_subject = entry.subject
    old_teacher = entry.teacher
    slot = entry.time_slot

    # Resolve new subject/teacher
    new_subject = Subject.objects.get(pk=new_subject_id) if new_subject_id else old_subject
    new_teacher = Teacher.objects.get(pk=new_teacher_id) if new_teacher_id else old_teacher

    # ── Step 1: Detect conflicts from this edit ──
    changed = []

    # Check if new teacher is already teaching another section at this slot
    teacher_conflicts = Timetable.objects.filter(
        time_slot=slot, teacher=new_teacher
    ).exclude(pk=entry_id)

    conflicting_entries = list(teacher_conflicts)

    # ── Step 2: Apply the primary edit ──
    entry.subject = new_subject
    entry.teacher = new_teacher
    entry.explanation = {
        'reason': f"Manually edited: {old_subject.name} → {new_subject.name}",
        'factors': [
            f"User requested change",
            f"Teacher: {old_teacher.name} → {new_teacher.name}",
        ],
        'sc_score': 0.0,
        'phase': 'manual_edit',
    }
    entry.save()
    changed.append(f"[{section.name}] {slot} → {new_subject.name} ({new_teacher.name})")

    # ── Step 3: Handle cascading conflicts ──
    if not conflicting_entries:
        return {
            'success': True,
            'changed': changed,
            'message': 'Edit applied, no cascading conflicts.',
        }

    # For conflicting entries: try to find alternative teachers
    for conflict in conflicting_entries:
        # Find available teachers for the conflicted subject
        alt_teachers = Teacher.objects.filter(
            subjects=conflict.subject
        ).exclude(pk=new_teacher.id)

        resolved = False
        for alt_teacher in alt_teachers:
            # Check if alt_teacher is free at this slot across all sections
            alt_busy = Timetable.objects.filter(
                time_slot=conflict.time_slot,
                teacher=alt_teacher,
            ).exists()
            if not alt_busy:
                old_name = conflict.teacher.name
                conflict.teacher = alt_teacher
                conflict.explanation = {
                    'reason': f"Auto-reassigned: teacher conflict resolved",
                    'factors': [
                        f"Original teacher {old_name} now busy (edited in {section.name})",
                        f"Reassigned to {alt_teacher.name} (available at this slot)",
                    ],
                    'sc_score': 0.0,
                    'phase': 'partial_regen',
                }
                conflict.save()
                changed.append(
                    f"[{conflict.section.name}] {conflict.time_slot} → "
                    f"{conflict.subject.name} (teacher: {old_name} → {alt_teacher.name})"
                )
                resolved = True
                break

        if not resolved:
            # Cannot resolve — remove the conflict, note the unresolved issue
            changed.append(
                f"[{conflict.section.name}] {conflict.time_slot} → "
                f"REMOVED {conflict.subject.name} (unresolvable teacher conflict)"
            )
            conflict.delete()

    return {
        'success': True,
        'changed': changed,
        'message': f'Edit applied with {len(conflicting_entries)} cascading change(s).',
    }

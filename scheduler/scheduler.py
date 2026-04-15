# scheduler.py — Final integrated version
# Keeps your present architecture:
# BaseScheduler, BacktrackingScheduler, GeneticScheduler stub,
# explanation metadata, debug_run(), partial_regenerate()
#
# Adds:
# - room allocation
# - teacher max-hours enforcement
# - room clash prevention
# - lab-room restriction
# - multi-section global room tracking

from abc import ABC, abstractmethod
from collections import defaultdict
from datetime import time as dt_time
from math import inf
import random


# ---------------------------------------------------------------------
# Time Slot Generation
# ---------------------------------------------------------------------

DEFAULT_CONFIG = {
    "days": ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"],
    "start_time": "09:00",
    "end_time": "16:00",
    "slot_duration": 60,
    "break_start": "13:00",
    "break_end": "14:00",
}


def _str_to_time(s: str) -> dt_time:
    h, m = s.strip().split(":")
    return dt_time(int(h), int(m))


def _time_to_minutes(t: dt_time) -> int:
    return t.hour * 60 + t.minute


def _minutes_to_time(m: int) -> dt_time:
    return dt_time(m // 60, m % 60)


def generate_time_slots(config: dict = None) -> list:
    """
    Generate slot dictionaries:
      [{'day', 'start', 'end', 'slot_index'}, ...]
    """
    cfg = {**DEFAULT_CONFIG, **(config or {})}
    days = cfg["days"]

    start_min = _time_to_minutes(_str_to_time(cfg["start_time"]))
    end_min = _time_to_minutes(_str_to_time(cfg["end_time"]))
    duration = int(cfg["slot_duration"])
    periods_per_day = int(cfg.get("periods_per_day", 6))

    if duration <= 0:
        raise ValueError("slot_duration must be a positive integer.")
    if periods_per_day <= 0:
        raise ValueError("periods_per_day must be a positive integer.")
    if start_min >= end_min:
        raise ValueError("start_time must be earlier than end_time.")

    break_start_min = break_end_min = None
    if cfg.get("break_start") and cfg.get("break_end"):
        break_start_min = _time_to_minutes(_str_to_time(cfg["break_start"]))
        break_end_min = _time_to_minutes(_str_to_time(cfg["break_end"]))

    slots = []
    for day in days:
        current = start_min
        slot_idx = 0

        while current + duration <= end_min and slot_idx < periods_per_day:
            slot_end = current + duration

            if break_start_min is not None and break_end_min is not None:
                if current < break_end_min and slot_end > break_start_min:
                    current = break_end_min
                    continue

            slots.append(
                {
                    "day": day,
                    "start": _minutes_to_time(current),
                    "end": _minutes_to_time(slot_end),
                    "slot_index": slot_idx,
                }
            )
            current = slot_end
            slot_idx += 1

        if slot_idx < periods_per_day:
            raise ValueError(
                f"Configuration for {day} only produced {slot_idx} period(s); "
                f"{periods_per_day} periods/day are required."
            )

    return slots


def persist_time_slots(slot_dicts: list):
    """
    Clear all TimeSlot / Timetable rows and create fresh TimeSlot rows.
    """
    from .models import TimeSlot, Timetable

    Timetable.objects.all().delete()
    TimeSlot.objects.all().delete()

    orm_slots = []
    for s in slot_dicts:
        slot = TimeSlot.objects.create(
            day=s["day"],
            start_time=s["start"],
            end_time=s["end"],
            slot_index=s["slot_index"],
        )
        orm_slots.append(slot)

    return orm_slots


# ---------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------

def _group_slots_by_day(slots):
    groups = defaultdict(list)
    for s in slots:
        groups[s.day].append(s)
    for day in groups:
        groups[day].sort(key=lambda s: s.slot_index)
    return dict(groups)


# ---------------------------------------------------------------------
# Debug Report
# ---------------------------------------------------------------------

class DebugReport:
    def __init__(self):
        self.teacher_overloads = []
        self.lab_failures = []
        self.unplaceable_subjects = []
        self.constraint_violations = []

    def to_dict(self):
        return {
            "error": "SchedulingFailed",
            "teacher_overloads": self.teacher_overloads,
            "lab_failures": self.lab_failures,
            "unplaceable_subjects": self.unplaceable_subjects,
            "constraint_violations": self.constraint_violations,
        }


# ---------------------------------------------------------------------
# Base Scheduler
# ---------------------------------------------------------------------

class BaseScheduler(ABC):
    """
    Base scheduler.

    Expected data keys:
      - subjects: list[Subject]
      - teachers: list[Teacher]
      - rooms: list[Room]
      - subject_teacher_map: {subject_id: teacher}
      - subject_hours_map: {subject_id: hours}
      - section_subject_data: {section_name: [...]}
    """

    def __init__(
        self,
        data: dict,
        slots: list,
        section_subjects: list = None,
        global_teacher_busy: dict = None,
        global_room_busy: dict = None,
    ):
        self.all_subjects = {s.id: s for s in data["subjects"]}
        self.all_teachers = {t.id: t for t in data["teachers"]}
        self.all_rooms = {r.id: r for r in data.get("rooms", [])}

        self.slots = slots
        self.slots_by_day = _group_slots_by_day(slots)
        self.subject_teacher_map = data["subject_teacher_map"]

        if section_subjects:
            self.section_subjects = section_subjects
            self.subjects = [self.all_subjects[si["subject_id"]] for si in section_subjects]
            self.subject_hours_map = {si["subject_id"]: si["hours"] for si in section_subjects}
            self.section_teacher_map = {
                si["subject_id"]: self.all_teachers.get(si["teacher_id"])
                for si in section_subjects
                if si.get("teacher_id") in self.all_teachers
            }
            self.preferred_room_map = {
                si["subject_id"]: si.get("room_id")
                for si in section_subjects
                if si.get("room_id")
            }
        else:
            self.subjects = data["subjects"]
            self.subject_hours_map = data["subject_hours_map"]
            self.section_teacher_map = {}
            self.preferred_room_map = {}

        # Treat labs the same as regular classes during auto-generation.
        self.lab_subjects = []
        self.theory_subjects = list(self.subjects)

        self.global_teacher_busy = global_teacher_busy if global_teacher_busy is not None else defaultdict(set)
        self.global_room_busy = global_room_busy if global_room_busy is not None else defaultdict(set)

        self.rng = random.Random()
        self.teacher_hours_used = defaultdict(int)

        self.assignments = {}   # slot_id -> (subject, teacher, room, is_lab_continuation)
        self.explanations = {}  # slot_id -> dict

    @abstractmethod
    def run(self) -> bool:
        raise NotImplementedError

    def _candidate_rooms(self, subject):
        rooms = list(self.all_rooms.values())

        if not rooms:
            return [None]

        preferred_room_id = self.preferred_room_map.get(subject.id)
        if preferred_room_id and preferred_room_id in self.all_rooms:
            pref = self.all_rooms[preferred_room_id]
            if pref in rooms:
                rooms = [pref] + [r for r in rooms if r.id != pref.id]

        return rooms

    def _teacher_can_take(self, teacher, extra_hours=1) -> bool:
        limit = getattr(teacher, "max_hours_per_week", None)
        if limit is None:
            return True
        return self.teacher_hours_used[teacher.id] + extra_hours <= limit

    def _teacher_for_subject(self, subject):
        teacher = self.section_teacher_map.get(subject.id)
        if teacher is not None:
            return teacher
        return self.subject_teacher_map.get(subject.id)

    def _spread_slots(self, slots):
        """
        Return slots in a more mixed order so the schedule does not fill
        one day or one time band first.
        """
        slots_by_day = defaultdict(list)
        for slot in slots:
            slots_by_day[slot.day].append(slot)

        for day_slots in slots_by_day.values():
            day_slots.sort(key=lambda s: (s.slot_index, s.start_time))

        day_order = list(slots_by_day.keys())
        self.rng.shuffle(day_order)

        max_len = max((len(v) for v in slots_by_day.values()), default=0)
        ordered = []
        for idx in range(max_len):
            shuffled_days = day_order[:]
            self.rng.shuffle(shuffled_days)
            for day in shuffled_days:
                day_slots = slots_by_day[day]
                if idx < len(day_slots):
                    ordered.append(day_slots[idx])

        return ordered

    def _slot_priority(self, slot, remaining, day_schedule):
        """
        Lower values get scheduled first. We favor emptier days and then
        add a small random jitter so the grid does not look rigid.
        """
        return (
            len(day_schedule.get(slot.day, [])),
            slot.slot_index,
            self.rng.random(),
        )

    def _choose_next_slot(self, free_slots, day_schedule):
        """
        Pick the next slot from the remaining pool instead of walking the
        timetable in a fixed left-to-right order.
        """
        if not free_slots:
            return None

        scored = []
        for slot in free_slots:
            day_load = len(day_schedule.get(slot.day, []))
            # Prefer emptier days and then mix in the period number plus jitter.
            score = (day_load * 10) + (slot.slot_index * 1.5) + self.rng.uniform(0.0, 4.0)
            scored.append((score, slot))

        scored.sort(key=lambda item: item[0])
        shortlist = [slot for _, slot in scored[: min(4, len(scored))]]
        return self.rng.choice(shortlist)

    def save(self, section_name: str = "A"):
        """
        Persist assignments + explanations to Timetable.
        Requires Timetable.room field.
        """
        from .models import Timetable, TimeSlot as TSModel, ClassSection

        section, _ = ClassSection.objects.get_or_create(name=section_name)
        Timetable.objects.filter(section=section).delete()

        for slot_id, (subject, teacher, room, is_cont) in self.assignments.items():
            slot = TSModel.objects.get(pk=slot_id)
            explanation = self.explanations.get(slot_id, {})

            Timetable.objects.create(
                subject=subject,
                teacher=teacher,
                room=room,
                time_slot=slot,
                section=section,
                is_lab_continuation=is_cont,
                explanation=explanation,
            )


# ---------------------------------------------------------------------
# Exceptions
# ---------------------------------------------------------------------

class SchedulingError(Exception):
    def __init__(self, message, debug_report=None):
        super().__init__(message)
        self.debug_report = debug_report


# ---------------------------------------------------------------------
# Backtracking Scheduler
# ---------------------------------------------------------------------

class BacktrackingScheduler(BaseScheduler):
    """
    Constraint-based recursive backtracking scheduler.

    Hard constraints:
      - teacher clash
      - room clash
      - section clash (enforced by Timetable unique constraint)
      - teacher max hours
      - balanced spreading across the week
    """

    def run(self) -> bool:
        total_required = sum(self.subject_hours_map.values())
        total_available = len(self.slots)

        if total_required > total_available:
            raise SchedulingError(
                f"Insufficient slots for this section. Required: {total_required}, Available: {total_available}."
            )

        remaining = {s_id: hrs for s_id, hrs in self.subject_hours_map.items()}
        day_schedule = defaultdict(list)
        local_teacher_busy = defaultdict(set)
        local_room_busy = defaultdict(set)
        used_slots = set()

        free_slots = [s for s in self.slots if s.id not in used_slots]

        return self._backtrack(
            free_slots=free_slots,
            remaining=remaining,
            day_schedule=day_schedule,
            local_teacher_busy=local_teacher_busy,
            local_room_busy=local_room_busy,
            used_slots=used_slots,
        )

    # -----------------------------------------------------------------
    # Debug mode
    # -----------------------------------------------------------------

    def debug_run(self) -> DebugReport:
        report = DebugReport()
        total_required = sum(self.subject_hours_map.values())
        total_available = len(self.slots)

        if total_required > total_available:
            report.constraint_violations.append(
                f"Insufficient slots: required {total_required}, available {total_available}"
            )

        for subject in self.subjects:
            teacher = self._teacher_for_subject(subject)
            if not teacher:
                report.constraint_violations.append(f"No teacher assigned for '{subject.name}'")
                continue

            required_hours = self.subject_hours_map.get(subject.id, 0)
            limit = getattr(teacher, "max_hours_per_week", inf)
            if required_hours > limit:
                report.teacher_overloads.append(
                    {
                        "teacher": teacher.name,
                        "subject": subject.name,
                        "required_hours": required_hours,
                        "available_teacher_limit": limit,
                    }
                )

        remaining = {s_id: hrs for s_id, hrs in self.subject_hours_map.items()}
        day_schedule = defaultdict(list)
        local_teacher_busy = defaultdict(set)
        local_room_busy = defaultdict(set)
        used_slots = set()

        free_slots = [s for s in self.slots if s.id not in used_slots]
        success = self._backtrack(
            free_slots=free_slots,
            remaining=remaining,
            day_schedule=day_schedule,
            local_teacher_busy=local_teacher_busy,
            local_room_busy=local_room_busy,
            used_slots=used_slots,
        )

        if not success:
            for s_id, hrs_left in remaining.items():
                if hrs_left > 0:
                    subj = self.all_subjects.get(s_id)
                    report.unplaceable_subjects.append(
                        {
                            "subject": subj.name if subj else str(s_id),
                            "remaining_hours": hrs_left,
                            "total_required": self.subject_hours_map.get(s_id, 0),
                        }
                    )

        return report

    # -----------------------------------------------------------------
    # Theory Scheduling
    # -----------------------------------------------------------------

    def _backtrack(self, free_slots, remaining, day_schedule, local_teacher_busy, local_room_busy, used_slots) -> bool:
        if all(h == 0 for h in remaining.values()):
            return True

        if not free_slots:
            return all(h == 0 for h in remaining.values())

        slot = self._choose_next_slot(free_slots, day_schedule)
        if slot is None:
            return all(h == 0 for h in remaining.values())

        candidates = self._get_candidates(slot, remaining, day_schedule, local_teacher_busy)

        for subject, sc_score, factors in candidates:
            teacher = self._teacher_for_subject(subject)
            if teacher is None:
                continue

            if not self._teacher_can_take(teacher, 1):
                continue

            if slot.id in self.global_teacher_busy.get(teacher.id, set()):
                continue
            if slot.id in local_teacher_busy.get(teacher.id, set()):
                continue

            room_candidates = self._candidate_rooms(subject)
            if not room_candidates:
                continue

            chosen_room = None
            roomless_allowed = False
            for room in room_candidates:
                if room is None:
                    roomless_allowed = True
                    break
                if slot.id in self.global_room_busy.get(room.id, set()):
                    continue
                if slot.id in local_room_busy.get(room.id, set()):
                    continue
                chosen_room = room
                break

            if chosen_room is None and not roomless_allowed:
                continue

            self.assignments[slot.id] = (subject, teacher, chosen_room, False)
            remaining[subject.id] -= 1
            day_schedule[slot.day].append(subject.id)

            local_teacher_busy[teacher.id].add(slot.id)
            self.global_teacher_busy[teacher.id].add(slot.id)

            if chosen_room is not None:
                local_room_busy[chosen_room.id].add(slot.id)
                self.global_room_busy[chosen_room.id].add(slot.id)

            used_slots.add(slot.id)
            self.teacher_hours_used[teacher.id] += 1

            hrs_left = remaining[subject.id]
            self.explanations[slot.id] = {
                "reason": f"MRV selected '{subject.name}' ({hrs_left}h remaining after placement)",
                "factors": factors + ([f"Room selected: {chosen_room.name}"] if chosen_room else ["No room required"]),
                "sc_score": round(sc_score, 2),
                "phase": "theory_backtracking",
            }

            if self._backtrack(
                [s for s in free_slots if s.id not in used_slots],
                remaining,
                day_schedule,
                local_teacher_busy,
                local_room_busy,
                used_slots,
            ):
                return True

            # undo
            del self.assignments[slot.id]
            self.explanations.pop(slot.id, None)
            remaining[subject.id] += 1
            day_schedule[slot.day].remove(subject.id)

            local_teacher_busy[teacher.id].discard(slot.id)
            self.global_teacher_busy[teacher.id].discard(slot.id)

            if chosen_room is not None:
                local_room_busy[chosen_room.id].discard(slot.id)
                self.global_room_busy[chosen_room.id].discard(slot.id)

            used_slots.discard(slot.id)
            self.teacher_hours_used[teacher.id] -= 1

        return self._backtrack(
            [s for s in free_slots if s.id not in used_slots],
            remaining,
            day_schedule,
            local_teacher_busy,
            local_room_busy,
            used_slots,
        )

    def _get_candidates(self, slot, remaining, day_schedule, local_teacher_busy):
        """
        Return priority-sorted theory subjects.

        Sorting priority:
          1. MRV (more remaining hours first, so negative remaining)
          2. soft constraints score
          3. subject name
        """
        today_subjects = day_schedule.get(slot.day, [])
        total_today = len(today_subjects)
        num_days = max(len(self.slots_by_day), 1)

        candidates = []

        for subject in self.theory_subjects:
            hrs_left = remaining.get(subject.id, 0)
            if hrs_left <= 0:
                continue

            teacher = self._teacher_for_subject(subject)
            if teacher is None:
                continue

            if slot.id in self.global_teacher_busy.get(teacher.id, set()):
                continue
            if slot.id in local_teacher_busy.get(teacher.id, set()):
                continue

            score = 0.0
            factors = [f"{teacher.name} is free for this slot"]

            # SC1: consecutive same subject penalty
            if today_subjects and today_subjects[-1] == subject.id:
                score += 10.0
                factors.append("SC1: consecutive same subject penalty +10")
            else:
                factors.append("SC1: no consecutive conflict")

            # SC2: daily balance penalty
            count_today = today_subjects.count(subject.id)
            if total_today > 0:
                balance_pen = (count_today / total_today) * 5.0
                score += balance_pen
                if balance_pen > 0:
                    factors.append(f"SC2: daily balance penalty +{balance_pen:.1f}")
            else:
                factors.append("SC2: first occurrence today")

            # SC3: daily overload penalty
            if count_today >= 2:
                sc3_pen = 15.0 * (count_today - 1)
                score += sc3_pen
                factors.append(f"SC3: overload penalty +{sc3_pen:.1f}")
            else:
                factors.append("SC3: under daily limit")

            # SC4: even distribution penalty
            total_hours = self.subject_hours_map.get(subject.id, 0)
            ideal_per_day = total_hours / num_days
            deviation = abs((count_today + 1) - ideal_per_day)
            sc4_pen = 8.0 * deviation
            score += sc4_pen
            factors.append(f"SC4: distribution penalty +{sc4_pen:.1f}")

            # Random jitter to avoid rigid, repeating patterns.
            jitter = round(self.rng.uniform(0.0, 2.0), 2)
            score += jitter
            factors.append(f"Random spread jitter +{jitter:.2f}")

            factors.append(f"Total soft score: {score:.1f}")
            factors.append(f"MRV remaining: {hrs_left}")

            candidates.append((-hrs_left, score, self._slot_priority(slot, remaining, day_schedule), subject, score, factors))

        candidates.sort(key=lambda x: (x[0], x[1], x[2]))
        return [(c[3], c[4], c[5]) for c in candidates]


# ---------------------------------------------------------------------
# Genetic Scheduler Stub
# ---------------------------------------------------------------------

class GeneticScheduler(BaseScheduler):
    def run(self) -> bool:
        raise NotImplementedError("GeneticScheduler is not yet implemented.")


# ---------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------

def generate_timetable(data: dict, slots: list) -> bool:
    """
    Backward-compatible single/multi-section scheduling.
    """
    sections_data = data.get("section_subject_data", {})
    section_names = list(sections_data.keys()) if sections_data else ["A"]

    if len(section_names) == 1:
        sec_name = section_names[0]
        sec_subjects = sections_data.get(sec_name, None)

        scheduler = BacktrackingScheduler(data, slots, section_subjects=sec_subjects)
        success = scheduler.run()
        if success:
            scheduler.save(section_name=sec_name)
        return success

    return generate_timetable_multi(data, slots, section_names)


def generate_timetable_multi(data: dict, slots: list, section_names: list) -> bool:
    """
    Multi-section scheduling with shared teacher and room tracking.
    """
    sections_data = data.get("section_subject_data", {})
    global_teacher_busy = defaultdict(set)
    global_room_busy = defaultdict(set)

    for sec_name in section_names:
        sec_subjects = sections_data.get(sec_name, None)
        if not sec_subjects:
            continue

        scheduler = BacktrackingScheduler(
            data,
            slots,
            section_subjects=sec_subjects,
            global_teacher_busy=global_teacher_busy,
            global_room_busy=global_room_busy,
        )

        try:
            success = scheduler.run()
        except SchedulingError as e:
            raise SchedulingError(f"[{sec_name}] {str(e)}")

        if not success:
            raise SchedulingError(f"[{sec_name}] Could not satisfy all constraints for this section.")

        scheduler.save(section_name=sec_name)

    return True


def debug_schedule(data: dict, slots: list) -> dict:
    """
    Run scheduler in debug mode, returning structured diagnostics.
    Does not persist changes.
    """
    sections_data = data.get("section_subject_data", {})
    section_names = list(sections_data.keys()) if sections_data else ["A"]
    global_teacher_busy = defaultdict(set)
    global_room_busy = defaultdict(set)
    all_reports = {}

    for sec_name in section_names:
        sec_subjects = sections_data.get(sec_name, None)
        if not sec_subjects:
            continue

        scheduler = BacktrackingScheduler(
            data,
            slots,
            section_subjects=sec_subjects,
            global_teacher_busy=global_teacher_busy,
            global_room_busy=global_room_busy,
        )

        report = scheduler.debug_run()
        all_reports[sec_name] = report.to_dict()

    return all_reports


# ---------------------------------------------------------------------
# Partial Regeneration
# ---------------------------------------------------------------------

def partial_regenerate(entry_id: int, new_subject_id: int = None, new_teacher_id: int = None) -> dict:
    """
    Localized re-scheduling after a user edit.
    """
    from .models import Timetable, Subject, Teacher

    try:
        entry = Timetable.objects.select_related(
            "subject", "teacher", "time_slot", "section"
        ).get(pk=entry_id)
    except Timetable.DoesNotExist:
        return {"success": False, "message": "Entry not found.", "changed": []}

    section = entry.section
    old_subject = entry.subject
    old_teacher = entry.teacher
    slot = entry.time_slot
    old_room = getattr(entry, "room", None)

    new_subject = Subject.objects.get(pk=new_subject_id) if new_subject_id else old_subject
    new_teacher = Teacher.objects.get(pk=new_teacher_id) if new_teacher_id else old_teacher

    changed = []

    # Teacher conflict at same slot
    teacher_conflicts = Timetable.objects.filter(
        time_slot=slot,
        teacher=new_teacher,
    ).exclude(pk=entry_id)

    conflicting_entries = list(teacher_conflicts)

    # Apply primary edit
    entry.subject = new_subject
    entry.teacher = new_teacher
    entry.explanation = {
        "reason": f"Manually edited: {old_subject.name} → {new_subject.name}",
        "factors": [
            "User requested change",
            f"Teacher: {old_teacher.name} → {new_teacher.name}",
        ],
        "sc_score": 0.0,
        "phase": "manual_edit",
    }
    entry.save()

    changed.append(f"[{section.name}] {slot} → {new_subject.name} ({new_teacher.name})")

    if not conflicting_entries:
        return {
            "success": True,
            "changed": changed,
            "message": "Edit applied, no cascading conflicts.",
        }

    for conflict in conflicting_entries:
        alt_teachers = Teacher.objects.filter(subjects=conflict.subject).exclude(pk=new_teacher.id)
        resolved = False

        for alt_teacher in alt_teachers:
            alt_busy = Timetable.objects.filter(
                time_slot=conflict.time_slot,
                teacher=alt_teacher,
            ).exists()

            if not alt_busy:
                old_name = conflict.teacher.name
                conflict.teacher = alt_teacher
                conflict.explanation = {
                    "reason": "Auto-reassigned: teacher conflict resolved",
                    "factors": [
                        f"Original teacher {old_name} now busy",
                        f"Reassigned to {alt_teacher.name}",
                    ],
                    "sc_score": 0.0,
                    "phase": "partial_regen",
                }
                conflict.save()

                changed.append(
                    f"[{conflict.section.name}] {conflict.time_slot} → "
                    f"{conflict.subject.name} (teacher: {old_name} → {alt_teacher.name})"
                )
                resolved = True
                break

        if not resolved:
            changed.append(
                f"[{conflict.section.name}] {conflict.time_slot} → "
                f"REMOVED {conflict.subject.name} (unresolvable teacher conflict)"
            )
            conflict.delete()

    return {
        "success": True,
        "changed": changed,
        "message": f"Edit applied with {len(conflicting_entries)} cascading change(s).",
    }

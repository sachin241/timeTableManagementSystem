# models.py — Database models for the AI Timetable Scheduler (v4)

from uuid import uuid4

from django.db import models
from django.db.models import Q


class ClassSection(models.Model):
    """Represents a class/branch such as CSE-A, CSE-B, ME-A, etc."""
    name = models.CharField(max_length=50, unique=True)

    class Meta:
        ordering = ["name"]

    def __str__(self):
        return self.name


class Subject(models.Model):
    """
    Academic subject.

    subject_type:
      - theory: standard single-slot scheduling
      - lab: consecutive multi-slot scheduling

    credits:
      - optional input
      - if hours_per_week is not set, save() can derive it from credits
    """
    SUBJECT_TYPES = [
        ("theory", "Theory"),
        ("lab", "Lab"),
    ]

    name = models.CharField(max_length=100, unique=True)
    hours_per_week = models.PositiveIntegerField(default=1)
    credits = models.PositiveIntegerField(
        null=True,
        blank=True,
        help_text="Optional. If provided without hours, auto-converts.",
    )
    subject_type = models.CharField(
        max_length=10,
        choices=SUBJECT_TYPES,
        default="theory",
    )

    def save(self, *args, **kwargs):
        if self.credits and (not self.hours_per_week or self.hours_per_week == 0):
            self.hours_per_week = self.credits * 2 if self.subject_type == "lab" else self.credits
        super().save(*args, **kwargs)

    def __str__(self):
        tag = " [LAB]" if self.subject_type == "lab" else ""
        return f"{self.name}{tag} ({self.hours_per_week}h/week)"


class Teacher(models.Model):
    """Represents a teacher who can teach one or more subjects."""
    name = models.CharField(max_length=100, unique=True)
    subjects = models.ManyToManyField(Subject, related_name="teachers")
    max_hours_per_week = models.PositiveIntegerField(default=20)

    def __str__(self):
        return self.name


class Room(models.Model):
    """Physical room or lab."""
    name = models.CharField(max_length=50, unique=True)
    is_lab = models.BooleanField(default=False)

    def __str__(self):
        return self.name


class TimeSlot(models.Model):
    """Represents a time period on a specific weekday."""
    DAYS_OF_WEEK = [
        ("Monday", "Monday"),
        ("Tuesday", "Tuesday"),
        ("Wednesday", "Wednesday"),
        ("Thursday", "Thursday"),
        ("Friday", "Friday"),
        ("Saturday", "Saturday"),
    ]

    day = models.CharField(max_length=10, choices=DAYS_OF_WEEK)
    start_time = models.TimeField()
    end_time = models.TimeField()
    slot_index = models.PositiveIntegerField(default=0)

    class Meta:
        ordering = ["day", "slot_index"]
        unique_together = [("day", "start_time", "end_time")]

    def __str__(self):
        return f"{self.day} {self.start_time.strftime('%H:%M')}–{self.end_time.strftime('%H:%M')}"


class Timetable(models.Model):
    """
    A single scheduled class entry.
    """
    subject = models.ForeignKey(Subject, on_delete=models.CASCADE)
    teacher = models.ForeignKey(Teacher, on_delete=models.CASCADE)
    room = models.ForeignKey(Room, on_delete=models.CASCADE, null=True, blank=True)
    time_slot = models.ForeignKey(TimeSlot, on_delete=models.CASCADE)
    section = models.ForeignKey(
        ClassSection,
        on_delete=models.CASCADE,
        related_name="timetable_entries",
    )

    is_lab_continuation = models.BooleanField(default=False)
    lab_group_id = models.UUIDField(default=uuid4, null=True, blank=True)
    explanation = models.JSONField(default=dict, blank=True)

    class Meta:
        ordering = ["section", "time_slot"]
        unique_together = [("time_slot", "section")]
        constraints = [
            models.UniqueConstraint(
                fields=["teacher", "time_slot"],
                name="unique_teacher_per_slot",
            ),
            models.UniqueConstraint(
                fields=["room", "time_slot"],
                condition=Q(room__isnull=False),
                name="unique_room_per_slot",
            ),
        ]

    def __str__(self):
        tag = " (cont.)" if self.is_lab_continuation else ""
        room_txt = f" @ {self.room.name}" if self.room else ""
        return (
            f"[{self.section.name}] {self.time_slot} → "
            f"{self.subject.name}{tag} ({self.teacher.name}){room_txt}"
        )
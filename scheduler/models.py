"""
models.py — Database models for the AI Timetable Scheduler (v4).

v4 Changes:
  - Timetable gains: explanation JSONField for per-slot reasoning metadata.

Models:
  - ClassSection: A class/branch (e.g. CSE-A).
  - Subject: Academic subject with hours/credits and theory/lab type.
  - Teacher: Teacher linked to subjects via M2M.
  - TimeSlot: Time period on a weekday (generated programmatically).
  - Timetable: Assignment of subject+teacher to slot+section + explanation.
"""

from django.db import models


class ClassSection(models.Model):
    """Represents a class/branch such as CSE-A, CSE-B, ME-A, etc."""
    name = models.CharField(max_length=50, unique=True)

    class Meta:
        ordering = ['name']

    def __str__(self):
        return self.name


class Subject(models.Model):
    """
    Represents an academic subject.

    Subject type determines scheduling behavior:
      - theory: Standard single-slot scheduling.
      - lab: Requires consecutive multi-slot blocks.

    Credits system (optional):
      - If credits is set and hours_per_week is 0, auto-calculate:
          theory: 1 credit = 1 hour
          lab:    1 credit = 2 hours
      - If both are set, hours_per_week takes priority.
    """
    SUBJECT_TYPES = [
        ('theory', 'Theory'),
        ('lab', 'Lab'),
    ]

    name = models.CharField(max_length=100, unique=True)
    hours_per_week = models.PositiveIntegerField(default=1)
    credits = models.PositiveIntegerField(null=True, blank=True,
                                          help_text="Optional. If provided without hours, auto-converts.")
    subject_type = models.CharField(max_length=10, choices=SUBJECT_TYPES, default='theory')

    def save(self, *args, **kwargs):
        """Auto-calculate hours_per_week from credits if not explicitly set."""
        if self.credits and (not self.hours_per_week or self.hours_per_week == 0):
            if self.subject_type == 'lab':
                self.hours_per_week = self.credits * 2
            else:
                self.hours_per_week = self.credits
        super().save(*args, **kwargs)

    def __str__(self):
        tag = ' [LAB]' if self.subject_type == 'lab' else ''
        return f"{self.name}{tag} ({self.hours_per_week}h/week)"


class Teacher(models.Model):
    """Represents a teacher who can teach one or more subjects."""
    name = models.CharField(max_length=100, unique=True)
    subjects = models.ManyToManyField(Subject, related_name='teachers')
    max_hours_per_week = models.IntegerField(default=20)
    def __str__(self):
        return self.name



class TimeSlot(models.Model):
    """
    Represents a time period on a specific weekday.
    Generated programmatically by scheduler.generate_time_slots().
    """
    DAYS_OF_WEEK = [
        ('Monday', 'Monday'),
        ('Tuesday', 'Tuesday'),
        ('Wednesday', 'Wednesday'),
        ('Thursday', 'Thursday'),
        ('Friday', 'Friday'),
        ('Saturday', 'Saturday'),
    ]

    day = models.CharField(max_length=10, choices=DAYS_OF_WEEK)
    start_time = models.TimeField()
    end_time = models.TimeField()
    slot_index = models.PositiveIntegerField(default=0)

    class Meta:
        ordering = ['day', 'slot_index']
        unique_together = [('day', 'start_time', 'end_time')]

    def __str__(self):
        return f"{self.day} {self.start_time.strftime('%H:%M')}–{self.end_time.strftime('%H:%M')}"

class Timetable(models.Model):
    subject = models.ForeignKey(Subject, on_delete=models.CASCADE)
    teacher = models.ForeignKey(Teacher, on_delete=models.CASCADE)
    room = models.ForeignKey('Room', on_delete=models.CASCADE, null=True)  # NEW

    time_slot = models.ForeignKey(TimeSlot, on_delete=models.CASCADE)
    section = models.ForeignKey(ClassSection, on_delete=models.CASCADE, related_name='timetable_entries')

    is_lab_continuation = models.BooleanField(default=False)
    lab_group_id = models.UUIDField(null=True, blank=True)  # OPTIONAL

    explanation = models.JSONField(default=dict, blank=True)

    class Meta:
        ordering = ['section', 'time_slot']
        unique_together = [('time_slot', 'section')]

        constraints = [
            models.UniqueConstraint(
                fields=['teacher', 'time_slot'],
                name='unique_teacher_per_slot'
            ),
            models.UniqueConstraint(
                fields=['room', 'time_slot'],
                name='unique_room_per_slot'
            )
        ]
# Added model for Rooms 

class Room(models.Model):
    name = models.CharField(max_length=50)
    is_lab = models.BooleanField(default=False)

    def __str__(self):
        return self.name
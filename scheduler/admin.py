"""Register models in Django admin for v4."""

from django.contrib import admin
from .models import Subject, Teacher, TimeSlot, Timetable, ClassSection


@admin.register(ClassSection)
class ClassSectionAdmin(admin.ModelAdmin):
    list_display = ('name',)
    search_fields = ('name',)


@admin.register(Subject)
class SubjectAdmin(admin.ModelAdmin):
    list_display = ('name', 'hours_per_week', 'credits', 'subject_type')
    list_filter = ('subject_type',)
    search_fields = ('name',)


@admin.register(Teacher)
class TeacherAdmin(admin.ModelAdmin):
    list_display = ('name',)
    filter_horizontal = ('subjects',)
    search_fields = ('name',)


@admin.register(TimeSlot)
class TimeSlotAdmin(admin.ModelAdmin):
    list_display = ('day', 'start_time', 'end_time', 'slot_index')
    list_filter = ('day',)
    ordering = ('day', 'slot_index')


@admin.register(Timetable)
class TimetableAdmin(admin.ModelAdmin):
    list_display = ('section', 'time_slot', 'subject', 'teacher', 'is_lab_continuation', 'explanation_preview')
    list_filter = ('section__name', 'time_slot__day', 'subject', 'is_lab_continuation')
    search_fields = ('subject__name', 'teacher__name', 'section__name')

    @admin.display(description='Explanation')
    def explanation_preview(self, obj):
        exp = obj.explanation or {}
        reason = exp.get('reason', '—')
        return reason[:80] + '…' if len(reason) > 80 else reason



from .models import Room

admin.site.register(Room)
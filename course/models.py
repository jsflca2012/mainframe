from django.db import models
from uuid import uuid4

from account.models import Instructor


class CourseCategory(models.Model):

    name = models.CharField(max_length=100)
    description = models.CharField(max_length=1000)


class Course(models.Model):
    CLASS = 'C'
    TUTORING = 'T'
    TYPE_CHOICES = (
        (CLASS, 'Class'),
        (TUTORING, 'Tutoring'),
    )

    # Course information
    type = models.CharField(
        max_length=1,
        choices=TYPE_CHOICES,
        default=CLASS,
    )
    subject = models.CharField(max_length=100)
    description = models.CharField(max_length=1000)
    instructor = models.ForeignKey(Instructor, on_delete=models.CASCADE)
    tuition = models.DecimalField(max_digits=6, decimal_places=2)

    # Logistical information
    room = models.CharField(max_length=50)
    days = models.CharField(max_length=10)
    schedule = models.TimeField()
    start_date = models.DateField()
    end_date = models.DateField()
    max_capacity = models.IntegerField()

    # Many-to-one relationship with CourseCategory
    course_category = models.ForeignKey(
        CourseCategory, on_delete=models.CASCADE, null=True)
from django.db import models
from course.models import CourseCategory

# Create your models here.


class Price(models.Model):
    # Basic price information
    name = models.CharField(
        max_length=1000,
        blank=True,
        null=True,
    )
    hourly_tuition = models.DecimalField(max_digits=5, decimal_places=2)

    class Meta:
        abstract = True


class PriceRule(Price):
    # Price Rule fields
    category = models.ForeignKey(
        CourseCategory, on_delete=models.PROTECT, null=True)
    ELEMENTARY_LVL = "ELEMENTARY_LVL"
    MIDDLE_LVL = "MIDDLE_LVL"
    HIGH_LVL = "HIGH_LVL"
    COLLEGE_LVL = "COLLEGE_LVL"

    ACADEMIC_CHOICES = (
        (ELEMENTARY_LVL, "Elementary"),
        (MIDDLE_LVL, "Middle"),
        (HIGH_LVL, "High"),
        (COLLEGE_LVL, "College"),
    )
    academic_level = models.CharField(
        max_length=10,
        choices=ACADEMIC_CHOICES,
        blank=True,
        null=True,
    )
    min_capacity = models.IntegerField(null=True, blank=True)
    max_capacity = models.IntegerField(null=True, blank=True)


class StaticPrice(Price):
    # List of associated courses with the tuition
    @property
    def course_list(self):
        return [course_id for course_id in self.course_set.all()]

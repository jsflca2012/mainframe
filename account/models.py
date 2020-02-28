from django.contrib.auth import get_user_model
from django.core.validators import MaxValueValidator, MinValueValidator
from django.db import models
from django_localflavor_us.us_states import US_STATES
from django.db.models import Q
from django.conf import settings


class UserInfo(models.Model):
    # Account type
    STUDENT_TYPE = 'student'
    PARENT_TYPE = 'parent'
    INSTRUCTOR_TYPE = 'instructor'
    ADMIN_TYPE = 'admin'
    ACCOUNT_TYPE_CHOICES = (
        (STUDENT_TYPE, 'Student'),
        (PARENT_TYPE, 'Parent'),
        (INSTRUCTOR_TYPE, 'Instructor'),
        (ADMIN_TYPE, 'Admin'),
    )

    # Gender
    MALE_GENDER = 'male'
    FEMALE_GENDER = 'female'
    UNSPECIFIED_GENDER = 'unspecified'
    GENDER_CHOICES = (
        (MALE_GENDER, 'Male'),
        (FEMALE_GENDER, 'Female'),
        (UNSPECIFIED_GENDER, 'Unspecified'),
    )
    STATE_CHOICES = tuple(sorted(US_STATES, key=lambda obj: obj[1]))

    user = models.OneToOneField(
        get_user_model(),
        on_delete=models.PROTECT,
        primary_key=True,
    )
    account_type = models.CharField(max_length=20, choices=ACCOUNT_TYPE_CHOICES)
    user_uuid = models.CharField(max_length=50, blank=True, null=True)
    gender = models.CharField(
        max_length=20,
        choices=GENDER_CHOICES,
        default=UNSPECIFIED_GENDER,
    )
    birth_date = models.DateField(blank=True, null=True)

    # Address
    address = models.CharField(max_length=64, blank=True, null=True)
    city = models.CharField(max_length=32, blank=True, null=True)
    phone_number = models.CharField(max_length=50, blank=True, null=True)
    state = models.CharField(max_length=16, choices=STATE_CHOICES, blank=True, null=True)
    zipcode = models.CharField(max_length=10, blank=True, null=True)

    # Timestamps
    updated_at = models.DateTimeField(auto_now=True)
    created_at = models.DateTimeField(auto_now_add=True)

    def __str__(self):
        return self.user.first_name + ' ' + self.user.last_name

    class Meta:
        abstract = True


class StudentManager(models.Manager):
    def search(self, query=None, qs_initial=None):
        if qs_initial is None or len(qs_initial) == 0:
            qs = self.get_queryset()
        else:
            qs = qs_initial

        if query is not None:
            or_lookup = (Q(user__first_name__icontains=query) |
                Q(user__last_name__icontains=query) |
                Q(user__email__iexact=query) |
                Q(user_uuid__iexact=query) |
                Q(address__icontains=query) |
                Q(city__icontains=query) |
                Q(phone_number__icontains=query) |
                Q(state__icontains=query) |
                Q(zipcode__icontains=query) |
                Q(school__icontains=query) |
                Q(primary_parent__user__first_name__icontains=query) |
                Q(primary_parent__user__last_name__icontains=query) |
                Q(secondary_parent__user__first_name__icontains=query) |
                Q(secondary_parent__user__last_name__icontains=query))
            try:
                query = int(query)
                or_lookup |= Q(grade=query)
            except ValueError:
                pass
            qs = qs.filter(or_lookup).distinct() # distinct() is often necessary with Q lookups
        return qs


class Note(models.Model):
    timestamp = models.DateTimeField(auto_now_add=True)
    title = models.TextField(blank=True)
    body = models.TextField()
    user = models.ForeignKey(
        settings.AUTH_USER_MODEL,
        on_delete=models.PROTECT,
    )
    important = models.BooleanField(default=False)
    complete = models.BooleanField(default=False)


class Student(UserInfo):
    # 0 is preschool/kindergarten, 13 is graduated
    grade = models.IntegerField(
        validators=[MinValueValidator(0), MaxValueValidator(13)],
        null=True,
        blank=True,
    )
    school = models.CharField(max_length=64, null=True, blank=True)

    primary_parent = models.ForeignKey(
        "Parent",
        on_delete=models.PROTECT,
        blank=True,
        null=True,
        related_name="student_primary_parent",
    )

    secondary_parent = models.ForeignKey(
        "Parent",
        on_delete=models.PROTECT,
        blank=True,
        null=True,
        related_name="student_secondary_parent",
    )

    objects = StudentManager()

    @property
    def enrollment_id_list(self):
        return [enrollment.id for enrollment in self.enrollment_set.all()]


class ParentManager(models.Manager):
    def search(self, query=None, qs_initial=None):
        if qs_initial is None or len(qs_initial) == 0:
            qs = self.get_queryset()
        else:
            qs = qs_initial

        if query is not None:
            or_lookup = (Q(user__first_name__icontains=query) |
                Q(user__last_name__icontains=query) |
                Q(user__email__icontains=query) |
                Q(address__icontains=query) |
                Q(city__icontains=query) |
                Q(phone_number__icontains=query) |
                Q(state__icontains=query) |
                Q(zipcode__icontains=query))
            qs = qs.filter(or_lookup).distinct()
        return qs


class Parent(UserInfo):
    MOTHER_REL = "mother"
    FATHER_REL = "father"
    GUARDIAN_REL = "guardian"
    OTHER_REL = "other"

    RELATIONSHIP_CHOICES = (
        (MOTHER_REL, "Mother"),
        (FATHER_REL, "Father"),
        (GUARDIAN_REL, "Guardian"),
        (OTHER_REL, "Other"),
    )
    relationship = models.CharField(
        max_length=20,
        choices=RELATIONSHIP_CHOICES,
        blank=True,
        null=True,
    )
    balance = models.DecimalField(decimal_places=2, default=0.0, max_digits=6)
    secondary_phone_number = models.CharField(max_length=50, blank=True, null=True)
    objects = ParentManager()

    @property
    def student_list(self):
        return [student.user.id for student in self.student_primary_parent.all().union(
            self.student_secondary_parent.all())]


class InstructorManager(models.Manager):
    def search(self, query=None, qs_initial=None):
        if qs_initial is None or len(qs_initial) == 0:
            qs = self.get_queryset()
        else:
            qs = qs_initial

        if query is not None:
            or_lookup = (Q(user__first_name__icontains=query) |
                Q(user__last_name__icontains=query) |
                Q(user__email__icontains=query) |
                Q(address__icontains=query) |
                Q(city__icontains=query) |
                Q(phone_number__icontains=query) |
                Q(state__icontains=query) |
                Q(zipcode__icontains=query))
            qs = qs.filter(or_lookup).distinct()
        return qs


class Instructor(UserInfo):
    objects = InstructorManager()

    biography = models.CharField(max_length=2000, null=True, blank=True)
    experience = models.CharField(max_length=2000, null=True, blank=True)
    language = models.CharField(max_length=2000, null=True, blank=True)
    subjects = models.CharField(max_length=2000, null=True, blank=True)


class InstructorAvailability(models.Model):
    instructor = models.ForeignKey(
        Instructor,
        on_delete=models.PROTECT,
    )

    DAYS_OF_WEEK = (
        ('monday', 'Monday'),
        ('tuesday', 'Tuesday'),
        ('wednesday', 'Wednesday'),
        ('thursday', 'Thursday'),
        ('friday', 'Friday'),
        ('saturday', 'Saturday'),
        ('sunday', 'Sunday'),
    )

    day_of_week = models.CharField(max_length=9, choices=DAYS_OF_WEEK)
    start_time = models.TimeField()
    end_time = models.TimeField()

    # Timestamps
    updated_at = models.DateTimeField(auto_now=True)
    created_at = models.DateTimeField(auto_now_add=True)


class InstructorOutOfOffice(models.Model):
    instructor = models.ForeignKey(
        Instructor,
        on_delete=models.PROTECT,
    )
    start_datetime = models.DateTimeField()
    end_datetime = models.DateTimeField()
    description = models.TextField(blank=True)
    all_day = models.BooleanField(default=False)

    # Timestamps
    updated_at = models.DateTimeField(auto_now=True)
    created_at = models.DateTimeField(auto_now_add=True)


class AdminManager(models.Manager):
    def search(self, query=None, qs_initial=None):
        if qs_initial is None or len(qs_initial) == 0:
            qs = self.get_queryset()
        else:
            qs = qs_initial

        if query is not None:
            or_lookup = (Q(user__first_name__icontains=query) |
                Q(user__last_name__icontains=query) |
                Q(user__email__icontains=query) |
                Q(address__icontains=query) |
                Q(city__icontains=query) |
                Q(phone_number__icontains=query) |
                Q(state__icontains=query) |
                Q(zipcode__icontains=query) |
                Q(admin_type__icontains=query))
            qs = qs.filter(or_lookup).distinct()
        return qs


class Admin(UserInfo):
    OWNER_TYPE = "owner"
    RECEPTIONIST_TYPE = "receptionist"
    ASSISSTANT_TYPE = "assisstant"

    TYPE_CHOICES = (
        (OWNER_TYPE, "Owner"),
        (RECEPTIONIST_TYPE, "Receptionist"),
        (ASSISSTANT_TYPE, "Assisstant"),
    )
    admin_type = models.CharField(
        max_length=20,
        choices=TYPE_CHOICES
    )

    objects = AdminManager()

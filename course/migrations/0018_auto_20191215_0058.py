# Generated by Django 2.2.4 on 2019-12-15 00:58

from django.db import migrations


class Migration(migrations.Migration):

    dependencies = [
        ('course', '0017_merge_20191209_0110'),
    ]

    operations = [
        migrations.RemoveField(
            model_name='coursenotecoursetag',
            name='course',
        ),
        migrations.RemoveField(
            model_name='coursenotecoursetag',
            name='note',
        ),
        migrations.DeleteModel(
            name='CourseNoteAccountTag',
        ),
        migrations.DeleteModel(
            name='CourseNoteCourseTag',
        ),
    ]

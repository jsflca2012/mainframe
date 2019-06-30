from django.contrib.auth import get_user_model
from account.models import (
    Student,
    Parent,
    Instructor
)
from rest_framework import serializers


class UserSerializer(serializers.ModelSerializer):
    class Meta:
        model = get_user_model()
        fields = ('username', 'password', 'email', 'first_name', 'last_name')


class StudentSerializer(serializers.ModelSerializer):
    user = UserSerializer()

    class Meta:
        model = Student
        read_only_fields = (
            'updated_at',
            'created_at',
            'user'
        )
        fields = (
            'gender',
            'address',
            'city',
            'phone_number',
            'state',
            'zipcode',
            'grade',
            'age',
            'school',
            'parent'
        )

    def create(self, validated_data):
        User = get_user_model()

        user_data = validated_data.pop('user')

        # new_user = User.objects.create(username=user_data.get('email'), password="asdf124124",
        #                                email=user_data.get('email'), first_name=user_data.get('first_name'), last_name=user_data.get('last_name'))

        # return new_user


class ParentSerializer(serializers.ModelSerializer):
    class Meta:
        model = Parent
        read_only_fields = (
            'updated_at',
            'created_at',
        )
        fields = (
            'user',
            'gender',
            'address',
            'city',
            'phone_number',
            'state',
            'zipcode',
            'relationship'
        )


class InstructorSerializer(serializers.ModelSerializer):
    class Meta:
        model = Instructor
        read_only_fields = (
            'updated_at',
            'created_at',
        )
        fields = (
            'user',
            'gender',
            'address',
            'city',
            'phone_number',
            'state',
            'zipcode',
            'age'
        )

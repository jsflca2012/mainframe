from itertools import chain
from rest_framework import generics
from rest_framework.response import Response

from account.models import (
    Student,
    Instructor
)

from course.models import (
    Course,
    Enrollment
)

from search.serializers import SearchViewSerializer

class AccountsSuggestionsView(generics.ListAPIView):
    serializer_class = SearchViewSerializer

    def get_queryset(self):
        #Student.objects.all()
        #Instructor.objects.all()
        return Student.objects.none()
        

class CoursesSuggestionsView(generics.ListAPIView):
    serializer_class = SearchViewSerializer
    
    def get_queryset(self):
        #Course.objects.all()
        return Course.objects.none()


class AccountsSearchView(generics.ListAPIView): 
    serializer_class = SearchViewSerializer

    def get_queryset(self):
        searchResults = Student.objects.none()

        # query input check
        query = self.request.query_params.get('query', None)
        if query is None or False in [word.isalnum() for word in query.split()]:
            return list(searchResults) 

        # query with profile filter
        profileFilter = self.request.query_params.get('profileFilter', None)
        if profileFilter is not None:
            if profileFilter == "Student":
                searchResults = Student.objects.search(query);

                gradeFilter = self.request.query_params.get('gradeFilter', None)
                if gradeFilter is not None:
                    try:
                        gradeFilter = int(gradeFilter)
                        if 1 <= gradeFilter and gradeFilter <= 12:
                            searchResults = searchResults.filter(grade=gradeFilter)
                    except ValueError:
                        pass
            elif profileFilter == "Instructor":
                searchResults = Instructor.objects.search(query)
            elif profileFilter == "Admin":
                searchResults = Admin.objects.search(query)
        # query on all models
        else:
            searchResults = chain(
                Student.objects.search(query),
                Parent.objects.search(query),
                Instructor.objects.search(query))

        # sort results
        sortFilter = self.request.query_params.get('sort', None)
        if sortFilter is not None:
            if sortFilter == "alphaAsc":
                searchResults = sorted(searchResults, key=lambda obj:obj.user.first_name)
            elif sortFilter == "alphaDesc":
                searchResults = sorted(searchResults, key=lambda obj:obj.user.first_name, reverse=True)
            elif sortFilter == "idAsc":
                searchResults = sorted(searchResults, key=lambda obj:obj.user.id)
            elif sortFilter == "idDesc": 
                searchResults = sorted(searchResults, key=lambda obj:obj.user.id, reverse=True)
        
        searchResults = list(searchResults)
        # extract page
        pageFilter = self.request.query_params.get('pageNumber', None)
        if pageFilter is not None:
            try:
                pageNumber = int(pageFilter)
                pageSize = 8
                resultLen = len(arr)
                rangeEnd = pageSize*pageNumber
                if pageNumber > 0 and rangeEnd-pageSize < resultLen:
                    searchResults = arr[rangeEnd-pageSize : resultLen if resultLen <= rangeEnd else rangeEnd]
            except ValueError:
                pass

        return searchResults

class CoursesSearchView(generics.ListAPIView):
    serializer_class = SearchViewSerializer

    def get_queryset(self):
        searchResults = Course.objects.none()

        # query input check
        query = self.request.query_params.get('query', None)
        if query is None or False in [word.isalnum() for word in query.split()]:
            return list(searchResults) 

        dateDic = {
            "monday":"MON",
            "tuesday":"TUE",
            "wednesday":"WED",
            "thursday":"THU",
            "friday":"FRI",
            "saturday":"SAT",
            "sunday":"SUN"
        }
        # date check
        if dateDic.get(query):
            searchResults = Course.objects.search(dateDic[query])
        else:
            searchResults = Course.objects.search(query)

        # course filter
        courseFilter = self.request.query_params.get('courseTypeFilter', None)
        if courseFilter is not None:
            if courseFilter == "tutoring":
                searchResults = searchResults.filter(max_capacity = 1)
            if courseFilter == "group":
                searchResults = searchResults.filter(max_capacity__lt = 5) 
            if courseFilter == "class":
                searchResults = searchResults.filter(max_capacity__gt = 5)

        # availability filter
        availabilityFilter = self.request.query_params.get('availability', None)
        if availabilityFilter is not None and (availabilityFilter == "open" or availabilityFilter == "filled"):
            # calculate availability
            course_ids = []
            for course in searchResults:
                capacity = len(Enrollment.objects.filter(course = course.course_id))
                if availabilityFilter == "open" and capacity < course.max_capacity:
                    course_ids.append(course.course_id)
                if availabilityFilter == "filled" and capacity >= course.max_capacity:
                    course_ids.append(course.course_id)
            searchResults = Course.objects.filter(course_id__in = course_ids)        
            
        # sort results
        sortFilter = self.request.query_params.get('sort', None)
        if sortFilter is not None:
            if sortFilter == "dateAsc":
                searchResults = searchResults.order_by("start_date")
            if sortFilter == "dateDesc":
                searchResults = searchResults.order_by("-start_date")
            if sortFilter == "timeAsc":
                searchResults = searchResults.order_by("start_time")
            if sortFilter == "timeDesc":
                searchResults = searchResults.order_by("-start_time")
        return list(searchResults)

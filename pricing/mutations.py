import graphene

from pricing.models import (
    PriceRule,
    Discount,
    MultiCourseDiscount,
    DateRangeDiscount,
    PaymentMethodDiscount
)
from pricing.schema import (
    PriceRuleType,
    DiscountType,
    MultiCourseDiscountType,
    DateRangeDiscountType,
    PaymentMethodDiscountType
)
from course.models import (
    CourseCategory
)

from graphql_jwt.decorators import login_required, staff_member_required


class AcademicLevelEnum(graphene.Enum):
    ELEMENTARY_LVL = "elementary_lvl"
    MIDDLE_LVL = "middle_lvl"
    HIGH_LVL = "high_lvl"
    COLLEGE_LVL = "college_lvl"


class CourseTypeEnum(graphene.Enum):
    TUTORING = "tutoring"
    SMALL_GROUP = "small_group"
    CLASS = "class"


class AmountTypeEnum(graphene.Enum):
    PERCENT = "percent"
    FIXED = "fixed"


class CreatePriceRule(graphene.Mutation):
    class Arguments:
        name = graphene.String(required=True)
        hourly_tuition = graphene.Float(required=True)
        category_id = graphene.Int(name="category", required=True)
        academic_level = AcademicLevelEnum(required=True)
        course_type = CourseTypeEnum(required=True)
    
    priceRule = graphene.Field(PriceRuleType)

    @staticmethod
    @staff_member_required
    def mutate(root, info, **validated_data):
        priceRule = PriceRule.objects.create(**validated_data)
        return CreatePriceRule(priceRule=priceRule)


class CreateDiscount(graphene.Mutation):
    class Arguments:
        name = graphene.String(required=True)
        description = graphene.String(required=True)
        amount = graphene.Float(required=True)
        amount_type = AmountTypeEnum(required=True)
        active = graphene.Boolean(required=True)

    discount = graphene.Field(DiscountType)

    @staticmethod
    @staff_member_required
    def mutate(root, info, **validated_data):
        discount = Discount.objects.create(**validated_data)
        return CreateDiscount(discount=discount)


class CreateMultiCourseDiscount(graphene.Mutation):
    class Arguments:
        name = graphene.String(required=True)
        description = graphene.String(required=True)
        amount = graphene.Float(required=True)
        amount_type = AmountTypeEnum(required=True)
        active = graphene.Boolean(required=True)
        num_sessions = graphene.Int(required=True)

    multiCourseDiscount = graphene.Field(MultiCourseDiscountType)

    @staticmethod
    @staff_member_required
    def mutate(root, info, **validated_data):
        multiCourseDiscount = MultiCourseDiscount.objects.create(**validated_data)
        return CreateMultiCourseDiscount(multiCourseDiscount=multiCourseDiscount)


class CreateDateRangeDiscount(graphene.Mutation):
    class Arguments:
        name = graphene.String(required=True)
        description = graphene.String(required=True)
        amount = graphene.Float(required=True)
        amount_type = AmountTypeEnum(required=True)
        active = graphene.Boolean(required=True)
        start_date = graphene.types.datetime.Date(required=True)
        end_date = graphene.types.datetime.Date(required=True)

    dateRangeDiscount = graphene.Field(DateRangeDiscountType)

    @staticmethod
    @staff_member_required
    def mutate(root, info, **validated_data):
        dateRangeDiscount = DateRangeDiscount.objects.create(**validated_data)
        return CreateDateRangeDiscount(dateRangeDiscount=dateRangeDiscount)


class CreatePaymentMethodDiscount(graphene.Mutation):
    class Arguments:
        name = graphene.String(required=True)
        description = graphene.String(required=True)
        amount = graphene.Float(required=True)
        amount_type = AmountTypeEnum(required=True)
        active = graphene.Boolean(required=True)
        payment_method =graphene.String(required=True)
    
    paymentMethodDiscount = graphene.Field(PaymentMethodDiscountType)

    @staticmethod
    @staff_member_required
    def mutate(root, info, **validated_data):
        paymentMethodDiscount = PaymentMethodDiscount.objects.create(**validated_data)
        return CreatePaymentMethodDiscount(paymentMethodDiscount=paymentMethodDiscount)


class Mutation(graphene.ObjectType):
    create_pricerule = CreatePriceRule.Field()
    create_discount = CreateDiscount.Field()
    create_multicoursediscount = CreateMultiCourseDiscount.Field()
    create_daterangediscount = CreateDateRangeDiscount.Field()
    create_paymentmethoddiscount = CreatePaymentMethodDiscount.Field()

from django.views.generic import ListView
from django.utils import timezone
from datetime import datetime
from operative.models import Resource, Workshop, StatisticsManager


class StatisticsView(ListView):
    template_name = 'table.html'
    context_object_name = 'statistics'

    def get_queryset(self):
        year = int(self.request.GET.get('year', timezone.now().year))
        month = int(self.request.GET.get('month', timezone.now().month))
        today = timezone.now().date()

        data = []

        for workshop in Workshop.objects.all():
            rows = []

            for resource in Resource.objects.all():
                month_stats = StatisticsManager.get_monthly_statistics(
                    workshop, resource, year, month
                )

                day_stats = StatisticsManager.get_daily_statistics(
                    workshop, resource, today
                )

                year_stats = StatisticsManager.get_yearly_statistics(
                    workshop, resource, year, current_month=month
                )


                rows.append({
                    'resource': resource,

                    'plan_month': month_stats['plan_month'],
                    'month_plan': month_stats['plan_to_date'],
                    'month_fact': month_stats['fact_month'],
                    'month_diff': month_stats['deviation'],

                    'day_plan': day_stats['plan_daily'],
                    'day_fact': day_stats['fact_daily'],
                    'day_diff': day_stats['deviation'],

                    'year_plan': year_stats['year_plan_to_date'],
                    'year_fact': year_stats['year_fact_to_date'],
                    'year_diff': year_stats['year_deviation'],
                })

            data.append({
                'workshop': workshop,
                'rows': rows
            })

        return data


    def get_context_data(self, **kwargs):
        context = super().get_context_data(**kwargs)
        context['year'] = int(self.request.GET.get('year', timezone.now().year))
        context['month'] = int(self.request.GET.get('month', timezone.now().month))
        return context

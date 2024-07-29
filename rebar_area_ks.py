class KoreanRebar:
    def __init__(self):
        self.rebar_areas = {
            0: 0,
            10: 71.33,
            13: 126.7,
            16: 198.6,
            19: 286.5,
            22: 387.1,
            25: 506.7,
            29: 642.4,
            32: 794.2,
            35: 956.6,
            38: 1140,
            41: 1340,
            51: 2027
        }

    def get_area(self, diameter):
        """
        주어진 직경에 대한 철근의 단면적을 반환합니다.
        
        :param diameter: 철근의 직경 (mm)
        :return: 철근의 단면적 (mm²)
        """
        if diameter in self.rebar_areas:
            return self.rebar_areas[diameter]
        else:
            raise ValueError(f"직경 {diameter}mm에 대한 정보가 없습니다.")

    def get_diameter_from_area(self, area, tolerance=1.0):
        """
        주어진 면적에 가장 가까운 철근의 직경을 반환합니다.
        
        :param area: 찾고자 하는 철근의 단면적 (mm²)
        :param tolerance: 허용 오차 (%)
        :return: 가장 가까운 철근의 직경 (mm)
        """
        closest_diameter = None
        min_difference = float('inf')

        for diameter, std_area in self.rebar_areas.items():
            difference = abs(std_area - area)
            if difference < min_difference:
                min_difference = difference
                closest_diameter = diameter

        if (min_difference / self.rebar_areas[closest_diameter]) * 100 <= tolerance:
            return closest_diameter
        else:
            raise ValueError(f"허용 오차 {tolerance}% 내의 면적 {area}mm²에 해당하는 철근을 찾을 수 없습니다.")

    def list_all_rebars(self):
        """
        모든 철근의 직경과 해당 단면적을 출력합니다.
        """
        for diameter, area in self.rebar_areas.items():
            print(f"직경 {diameter}mm: 단면적 {area}mm²")
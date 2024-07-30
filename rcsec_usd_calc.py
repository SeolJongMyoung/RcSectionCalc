import math
import numpy as np
from scipy import interpolate #직선보간법
from excel_data_processor import DataProcessor
from rebar_area_ks import KoreanRebar
from civil_usd_materials import *
#--------------------------------------------------
#    직사각형 단철근보 단면검토 (도로교설계기준 2010)
#--------------------------------------------------
class CalcReinfoeceConcrete:

    def __init__(self, excel_file_path):
        self.processor = DataProcessor(excel_file_path)
        self.processor.load_and_process_data()
        processed_data = self.processor.get_processed_data()
        self.rebar = KoreanRebar()
        
        self.f_ck = processed_data[0][0]           #fck설정
        self.f_y = processed_data[0][1]            #fy설정

        self.con_material = ConcMaterial(f_ck=self.f_ck)
        self.rebar_material = RebarMaterial(f_y=self.f_y)

        self.pi_f = processed_data[0][2]            #Øf설정
        self.pi_v = processed_data[0][3]            #Øv설정
        self.Mu = processed_data[0][4]            #Mu설정
        self.Vu = processed_data[0][5]            #Vu설정  
        self.Nu = processed_data[0][6]            #Nu설정
        self.Ms = processed_data[0][7]            #Ms설정
        self.beam_h = processed_data[0][8]             #단면두께 설정
        self.beam_b = processed_data[0][9]             #단면 폭 설정
        self.as_dia1 = processed_data[1][0]        #1단 철근직경 설정
        self.as_num1 = processed_data[1][1]        #1단 철근개수 설정
        self.dc_1 = processed_data[1][2]           #1단 피복두께 설정
        self.as_dia2 = processed_data[1][3]        #2단 철근직경 설정
        self.as_num2 = processed_data[1][4]        #2단 철근개수 설정
        self.dc_2 = processed_data[1][5]           #2단 피복두께 설정
        self.as_dia3 = processed_data[1][6]        #3단 철근직경 설정
        self.as_num3 = processed_data[1][7]        #3단 철근개수 설정
        self.dc_3 = processed_data[1][8]           #3단 피복두께 설정
        self.av_dia = processed_data[2][0]         #전단철근 직경
        self.av_leg = processed_data[2][1]         #전단철근 다리개수
        self.av_space = processed_data[2][2]       #전단철근 배치간격
        self.sg = processed_data[2][3]             #복부스트럿 각도 입력방법 선택
        self.seta = processed_data[2][4]           #복부스트럿 각도 직접입력값
        self.alpha_dgree = processed_data[2][5]    #전단철근과 주철근 각도(주철근으로부터 시계방향각도)
#        self.δ = 1                                #재분배 모멘트율 설정
        self.E_s = self.rebar_material.E_s   #철근의 탄성계수
        self.Mu_nm = self.Mu*1000000
        self.Vu_n = self.Vu*1000
        self.Nu_n = self.Nu*1000
#        self.αcc = 0.85                           #유효계수
        if self.f_y <= 300 :
            self.rebar_id = "D"
        else :
            self.rebar_id = "H"
        
        self.as_use1 = self.rebar.get_area(self.as_dia1)
        self.as_use2 = self.rebar.get_area(self.as_dia2)
        self.as_use3 = self.rebar.get_area(self.as_dia3)
        
        self.d_c = (self.as_use1*self.as_num1*self.dc_1 + self.as_use2*self.as_num2*self.dc_2 + self.as_use3*self.as_num3*self.dc_3)/(self.as_use1*self.as_num1 + self.as_use2*self.as_num2 + self.as_use3*self.as_num3)

#------------------------------
#          휨모멘트 검토
#------------------------------
    def calculate(self):
        # Perform all necessary calculations here
        self.calc_moment()
        # Add other calculation methods as needed
        return self.get_results()
    
    
    def calc_moment(self) :
        if self.f_ck <= 28 :
            self.beta_1 = 0.85
        else : 
            self.beta_1 = max( 0.65, 0.85 - ( self.f_ck - 28 ) * 0.007)
    
        self.lo_bal = (0.85 * self.beta_1 * self.f_ck / self.f_y) * (6000 / (6000 + self.f_y))                         #균형철근비
        self.lo_max = 0.75 * self.lo_bal                                                                               #최대철근비 
        self.lo_min_1 = 1.4 / self.f_y                                                                                 #최소철근비
        self.lo_min_2 = 0.8 * math.sqrt(self.f_ck) / self.f_y
                                                                                    
        self.lo_min = max(self.lo_min_1, self.lo_min_2)                                       

        self.d_eff = self.beam_h - self.d_c                                                                            #단면 유효높이
        self.as_use = self.as_use1 * self.as_num1 + self.as_use2 * self.as_num2 + self.as_use3 * self.as_num3          #전체 사용철근량
        self.lo_use = self.as_use/(self.beam_b*self.d_eff)                                                             #사용철근비

        self.tension_force = self.as_use * self.f_y                                                                    #단면의 총 인장력
        self.compression_force = 0.85 * self.f_ck * self.beam_b                                                        #단면의 종 압축력 / a
        self.a = self.tension_force / self.compression_force                                                           #등가응력블럭깊이
        self.c = self.a / self.beta_1
        self.epsilon_y = self.f_y / self.E_s
        self.epsilon_t = 0.003 * ( self.d_eff - self.c ) / self.c

        if self.epsilon_t >= 0.005 :
            self.epsilon_t_result = "인장지배단면"
            self.pi_f_r = self.pi_f
        elif self.epsilon_t > self.epsilon_y :
            self.epsilon_t_result = "압축지배단면"
            self.pi_f_r = (self.pi_f - 0.65 )/(0.005-self.epsilon_y) * (self.epsilon_t - 0.005) + self.pi_f
        else :
            self.epsilon_t_result = "압축지배단면"
            self.pi_f_r = 0.65

        temp_a = (self.f_y**2)/(2 * 0.85 * self.f_ck * self.beam_b)
        temp_b = -self.f_y * self.d_eff
        self.as_req = (-temp_b - (temp_b**2 - 4 * temp_a * self.Mu_nm)**(1/2)) / (2 * temp_a)                          #필요철근량 산정

        self.lo_min_3 = self.as_req * 4 / ( 3 * self.beam_b*self.d_eff )
        self.lo_min_f = max(self.lo_min, self.lo_min_3)

        as_min_1 = self.lo_min * self.beam_b * self.d_eff
        as_min_2 = self.lo_min_f * self.beam_b * self.d_eff
        
        self.as_max = self.lo_max * self.beam_b * self.d_eff

        self.M_r = self.as_use * self.pi_f_r * self.f_y * ( self.d_eff - self.a / 2 )
        
        self.M_sf = (self.M_r / self.Mu_nm)

#------------------------------
#          전단력 검토
#------------------------------

#    def calshear(self) :
#        self.k = 1+math.sqrt(200/self.D)     #단면크기효과 고려한 계수
#        if self.k > 2 :
#            self.k1 = 2
#        else :
#            self.k1 = self.k
#        self.fctk = 0.7*0.3*self.fcm**(2/3)     #콘크리트 인장강도
#        self.fnn = self.Nu*1000/(self.H*self.B)  #전단철근이 없는경우 축인장응력
#        self.fnmax = 0.2*self.Øc*self.fck       #이값이상 사용금지인지 단면을 증가시켜야 되는지 알수 없음
#        self.fn = min(self.fnn, self.fnmax)
#        self.ρs = min(self.ρ, 0.02)
#        self.Vc  = (0.85*self.Øc*self.k*(self.ρs*self.fck)**(1/3) + 0.15*self.fn)*(self.B*self.D)  #전단철근이 없는 부재의 설계전단강도
#       self.Vcdmin = (0.4*self.Øc*self.fctk + 0.15*self.fn)*(self.B*self.D)   #최소설계 전단강도
#        self.Vcd = max(self.Vc,self.Vcdmin)
#                    
#        self.Avs = self.rebar(self.AvDia)*self.AvLeg #전단철근량  
#        self.α = 90.0
#        self.ν = 0.6*(1 - self.fck/250)
#        self.z = 0.9*self.D
#        if self.fnn < 0 :
#           self.αcw = 0
#        elif self.fnn == 0:
#            self.αcw = 1.0
#        elif self.fnn <= 0.25*self.Øc*self.fck:
#            self.αcw = 1.0+self.fnn/(self.Øc*self.fck)
#       elif self.fnn <= 0.5*self.Øc*self.fck:
#            self.αcw = 1.25
#        elif self.fnn <= 1.0*self.Øc*self.fck:
#            self.αcw = 2.5*(1-self.fnn/(self.Øc*self.fck))    
#        else:
#            self.αcw = 0
#    
#        self.cotθ1 = 2.5              #cotθ = 2.5(θ=21.8도) 적용시
#        self.tanθ1 = 1/self.cotθ1
#        self.cotθ2 = 1                #cotθ = 1.0(θ=45.0도) 적용시
#        self.tanθ2 = 1/self.cotθ2
#        self.Vdmax1 = (self.ν*self.Øc*self.fck*self.B*self.z) / (self.cotθ1+self.tanθ1)
#        self.Vdmax2 = (self.ν*self.Øc*self.fck*self.B*self.z) / (self.cotθ2+self.tanθ2)
#        if self.sg == 1 :             #input에 1번 직접입력시 입력값 적용
#            self.cotθ = self.seta
#        elif self.sg == 2 :           #input에 2번 중간값입력시 중간값 적용  
#            self.cotθ = (1+2.5)/2
#        else :                        #input에 3번 자동산출입력시 자동계산 적용 (Eurocode 적용)
#            if self.Vun <= self.Vdmax1 :   #Vu가 Vmax1 보다 작은경우 cotθ = 2.5(θ=21.8도) 적용 
#                self.cotθ = 2.5
#            elif self.Vun > self.Vdmax2 :  #Vu가 Vmax2 보다 큰경우 cotθ = 0 적용으로 단면 증가 필요
#                self.cotθ = 0
#            else :
#                self.cotθ = 1/math.tan(0.5*math.asin(self.Vun /(0.2*self.fck*(1-self.fck/250)*self.B*self.z)))  #Vu가 Vmax2 보다 큰경우 산정불가식
#        self.tanθ = 1/self.cotθ
#        self.θ = math.degrees(math.atan(self.tanθ))
#                
#        self.Vd = (self.Øs*self.fy*self.Avs*0.9*self.D / self.AvSpace)*self.cotθ 
#        self.Vdmax = (self.ν*self.Øc*self.fck*self.B*self.z) / (self.cotθ+self.tanθ)#
#
#        self.ρvuse = self.Avs / (self.AvSpace*self.B*math.sin(self.α))
#        self.ρvmin = 0.08*math.sqrt(self.fck) / self.fy
#        self.s1max = 0.75*self.D*(1+(1/math.tan(self.α*math.pi/180)))   #종방향 전단철근 간격규정
#        self.s2max = min(0.75*self.D, 600)                              #횡방향 철근 최대폭 
#        self.s2 = self.B-2*self.Dc                                      #횡방향 철근 간격
#    

#--------------------------------------
#          사용성 검토(균열검토)
#--------------------------------------
#    def calservice(self):
#        self.nr  = round( self.Es / (0.077*(2300)**(1.5)*(self.fck + self.Δf)**(1/3)))     #철근비 산정(반올림)
#        self.Xo = (self.B*self.H**2/2 + (self.nr-1)*self.Asuse*self.D) / (self.B*self.H + (self.nr-1)*self.Asuse) 
#        self.Io = self.B*self.H**3/12 + self.B*self.H*(self.H/2-self.Xo)**2 + (self.nr-1)*self.Asuse*(self.D-self.Xo)**2
#        self.fct = self.Ms*10**6 / self.Io*(self.H-self.Xo)

#        self.fs = self.nr*self.Ms*10**6 / self.Io*(self.D-self.Xo) #사용철근의 응력

#        if self.fct <= self.fctk :
#            self.cr1 = "≤"
#        else:
#            self.cr1 = ">"

#        if self.fs < 0.8*self.fy:
#            self.cr3 = "≤"
#        else:
#            self.cr3 = ">"    
#        if self.fs < 0.8*self.fy:
#            self.cr4 = "... ∴O.K"
#        else:
#            self.cr4 = "... ∴N.G"
#
#        if self.fct <= self.fctk :
#            self.cr2 = "∴ 비균열 단면 ⇒ 균열검토 생략"
#        else:
#            self.cr2 = "∴ 균열 단면 검토"
#            self.k = -self.nr*self.ρ+ math.sqrt((self.nr*self.ρ)**2+ 2*self.nr*self.ρ)       #중립축비
#            self.x = self.k*self.D
#            self.fc = 2*self.Ms*10**6 / (self.B*self.x*(self.D - self.x/3))
#            self.fs = self.Ms*10**6 / (self.Asuse*(self.D - self.x/3))
#    
#            self.fsa = max(160,360)

#            if self.fsa >= self.fs:
#                self.cr5 = "≥"
#            else:
#                self.cr5 = "<"    
#            if self.fsa >= self.fs:
#                self.cr6 = "... ∴O.K"
#            else:
#                self.cr6 = "... ∴N.G"
#            self.Act = self.B*(self.H-self.Xo)
#            if self.Nu >= 0:       #압축(+)
#                self.k1 = 1.5
#            elif self.Nu < 0:      #인장(-)
#                if self.H < 1000:
#                    self.h1 = self.H
#                elif self.H >= 1000:
#                    self.h1 = 1000
#                self.k1 = 2*self.h1/(3*self.H)

#            if self.H <= 300:    #부등분포 반영 계수
#                self.k = 1.0
#            elif self.H <= 300 and self.H < 800:
#                self.k = -0.0007*self.H+1.21
#            elif self.H >= 800:
#                self.k = 0.65
#            self.fct = self.fctk / 0.7 #fct = fctm

#            if self.H < 1000:
#                self.h1 = self.H
#            elif self.H >= 1000:
#                self.h1 = 1000
        
#            self.kc = 0.4*(1 - self.fn/(self.k1*(self.H/self.h1)*self.fct))

#            self.Asdmin = self.kc*self.k*self.Act*self.fct / self.fsa   #최소철근량 산정
#            if self.Asdmin <= self.Asuse:
#                self.cr7 = "≤"
#            else:
#                self.cr7 = ">"    
#            if self.Asdmin <= self.Asuse:
#                self.cr8 = "... ∴O.K"
#            else:
#                self.cr8 = "... ∴N.G"
#        
#        print('새로 생성된 "section_check.txt"를 확인하세요!!')
#
    def get_results(self):
        # Return a dictionary or some structure with all the calculated results
        return {
            "f_ck": self.f_ck,
            "f_y": self.f_y,
            "as_dia1" : self.as_dia1,
            "as_dia2" : self.as_dia2,
            "as_dia3" : self.as_dia3,
            "use_as"  : self.as_use,
            "d"       : self.d_eff,
            "beta1"   : self.beta_1,
            "a"       : self.a,
            "ep_t"    : self.epsilon_t,
            "ep_y"    : self.epsilon_y, 
            "pi_f"    : self.pi_f_r,
            "pi_Mn"   : self.M_r,
            # Add other results as needed
        }

def main():
    file_path = "D:\Python\RC_Section_Calculation/Calc_As_input.xlsx"
    calculator = CalcReinfoeceConcrete(file_path)
    result = calculator.calculate()
    print(result)
    
if __name__ == "__main__":
    main()
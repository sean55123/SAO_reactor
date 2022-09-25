import win32com.client as win32
import os
import numpy as np
import glob
import time
import matplotlib.pyplot as plt
import random
import csv

class Aspen():
    def __init__(self):
        self.file = 'heatexchanging.apwz'
        self.file = os.path.abspath(self.file)
        self.folder_path = os.path.dirname(self.file)
        self.aspen = self.initiation()

    def delete_file(self):
        name = ['_*', '$*', 'aspenplus_processdump*', '*$back*']
        for i in name:
            path = self.folder_path + '/' + i
            del_file = glob.glob(path)
            for j in del_file:
                os.remove(j)

    def record(self, data_value, data_label):
        file_path = self.folder_path + '/reactor_data.csv'
        if os.path.isfile(file_path):
            with open('reactor_data.csv', 'a+', newline = '') as csvfile:
                writer = csv.writer(csvfile)
                writer.writerow(data_value)
        else:
            with open('reactor_data.csv', 'a+', newline = '') as csvfile:
                writer = csv.writer(csvfile)
                writer.writerow(data_label)
                writer.writerow(data_value)

    def initiation(self):
        self.delete_file()
        aspen = win32.Dispatch('Apwn.Document.37.0')
        aspen.InitFromFile2(self.file)
        aspen.Visible = 1
        aspen.SuppressDialogs = 1
        return aspen

    def check_status(self):
        total_status = 0
        for i in self.status_code:
            if (i & 1) == 1: #converge
                status = 1
            elif (i & 4) == 4: #error
                status = 4
            elif (i & 32) == 32: #warning
                status = 32
            else:
                status = 0 #else
            total_status += status
        return total_status

    def check_validation(self):
        self.status_code = []
        self.status_code.append(self.aspen.Tree.FindNode(r"\Data\Blocks\B1").AttributeValue(12))
        self.status_code.append(self.aspen.Tree.FindNode(r"\Data\Blocks\B2").AttributeValue(12))
        self.status_code.append(self.aspen.Tree.FindNode(r"\Data\Blocks\B3").AttributeValue(12))
        status = self.check_status()
        if status == 3:
            print("Converged!")
            return status
        else:
            print("Not converged :<")
            status = 404
            return status

    def diversed(self):
        self.aspen.Close()
        self.aspen = self.initiation()

    def get_original_params(self):
        """
        Variables : Feed temperature
                    Coolant flowrate
                    Coolant Pressure(set some to choose)
                    Coolant outlet from cooler temperature
                    reactor length B1, B2 >> Pressure drop may be in consideration
                    !!Bed voidage set to be 0.7!! 
        """
        feed_T = self.aspen.Tree.FindNode(r"\Data\Streams\S1\Output\RES_TEMP").value
        cool_F = self.aspen.Tree.FindNode(r"\Data\Streams\S8\Output\MASSFLMX\MIXED").value
        cool_P = self.aspen.Tree.FindNode(r"\Data\Streams\S8\Output\PRES_OUT\MIXED").value
        cool_OT = self.aspen.Tree.FindNode(r"\Data\Blocks\B3\Output\B_TEMP").value
        r_L1 = self.aspen.Tree.FindNode(r"\Data\Blocks\B1\Output\LEN_REACTOR").value
        r_D1 = self.aspen.Tree.FindNode(r"\Data\Blocks\B1\Output\DIAMETER").value
        r_L2 = self.aspen.Tree.FindNode(r"\Data\Blocks\B2\Output\LEN_REACTOR").value
        r_D2 = self.aspen.Tree.FindNode(r"\Data\Blocks\B2\Output\DIAMETER").value
        variables_list = [feed_T, cool_F, cool_OT, r_L1, r_D1, r_L2, r_D2, cool_P]
        return variables_list

    def set_newparams(self, current_solution):
        self.aspen.Tree.FindNode(r"\Data\Streams\S1\Input\TEMP").value == current_solution[0]
        self.aspen.Tree.FindNode(r"\Data\Streams\S8\Input\TOTFLOW").value == current_solution[1]
        self.aspen.Tree.FindNode(r"\Data\Streams\S8\Input\PRES").value == current_solution[7]
        self.aspen.Tree.FindNode(r"\Data\Blocks\B3\Input\TEMP").value == current_solution[2]
        self.aspen.Tree.FindNode(r"\Data\Blocks\B1\Input\LENGTH").value == current_solution[3]
        self.aspen.Tree.FindNode(r"\Data\Blocks\B1\Input\DIAM").value == current_solution[4]
        self.aspen.Tree.FindNode(r"\Data\Blocks\B2\Input\LENGTH").value == current_solution[5]
        self.aspen.Tree.FindNode(r"\Data\Blocks\B2\Input\DIAM").value == current_solution[6]
        self.aspen.Tree.FindNode(r"\Data\Blocks\B3\Input\PRES").value == (current_solution[7] - 0.3)

        self.pressure_drop(current_solution, reactor_num=1)
        self.pressure_drop(current_solution, reactor_num=2)
        time.sleep(2)
        self.aspen.Run2()
        print('Excustion complete!')
        time.sleep(2)
    
    def get_result(self):
        """
        First try CO2 conversion
        ! make sure the L/D ratio no larger than 5
        ! make sure temperature no larger than 400
        """
        inlet = self.aspen.Tree.FindNode(r"\Data\Streams\S1\Output\MOLEFLOW\MIXED\CO2").value
        outlet_1 = self.aspen.Tree.FindNode(r"\Data\Streams\S5\Output\MOLEFLOW\MIXED\CO2").value
        outlet_2 = self.aspen.Tree.FindNode(r"\Data\Streams\S2\Output\MOLEFLOW\MIXED\CO2").value
        for i in range(2):
            bname = ['B1', 'B2']
            for j in range(50):
                Tp = self.aspen.Tree.FindNode("\\Data\Blocks\\" + bname[i] + "\\Output\\B_TEMP2\\PROCESS\\" + str(j+1)).value
                Tc = self.aspen.Tree.FindNode("\\Data\Blocks\\" + bname[i] + "\Output\B_TEMP2\COOLANT\1" + str(j+1)).value
                if Tp >= 400:
                    overshoot = True
                else:
                    overshoot = False
                if (Tp - Tc) < 5:
                    close = True
                else:
                    close = False

        conversion = [] # [1, 2, overall]
        conversion.append((inlet - outlet_1) / inlet)
        conversion.append((outlet_1 - outlet_2) / outlet_1)
        conversion.append(conversion[0] + conversion[1])
        return conversion, overshoot, close

    def pressure_drop(self, current_solution, reactor_num):
        Dp = 3.6e-3 # m, Assumed to be 3.6 mm
        mu = 2.71e-5 # kg/ms
        phi = 0.7
        if reactor_num == 1:            
            Ac = current_solution[4]**2 * np.pi / 4 # m**2
            MassF = self.aspen.Tree.FindNode(r"\Data\Streams\S1\Output\MASSFLMX\MIXED").value # kg/hr
            MassF /= 3600 # kg/sec
            V = self.aspen.Tree.FindNode(r"\Data\Streams\S1\Output\VOLFLMX\MIXED").value # l/min
            V /= 6e4 # m3/sec 
            beta = V*(1-phi)/(Ac*Dp*phi**3) * (150*(1-phi)*mu/Dp + 1.75*MassF/Ac)
            p0 = 10 * 1e5 # bar to pa
            pd = p0 * (1 - 2*beta*current_solution[3]/p0)**0.5
            pd /= 1e5 # pa to bar
            pd -= p0 / 1e5
            self.pd1 = -pd
            self.aspen.Tree.FindNode(r"\Data\Blocks\B1\Input\PDROP").value = self.pd1
            print('Pressure Checked!')

        elif reactor_num == 2:
            Ac = current_solution[6]**2 * np.pi / 4 # m**2
            MassF = self.aspen.Tree.FindNode(r"\Data\Streams\S5\Output\MASSFLMX\MIXED").value # kg/hr
            MassF /= 3600 # kg/sec
            V = self.aspen.Tree.FindNode(r"\Data\Streams\S5\Output\VOLFLMX\MIXED").value # l/min
            V /= 6e4 # m3/sec 
            beta = V*(1-phi)/(Ac*Dp*phi**3) * (150*(1-phi)*mu/Dp + 1.75*MassF/Ac)
            p0 = (10 - self.pd1) * 1e5 # bar to pa
            pd = p0 * (1 - 2*beta*current_solution[5]/p0)**0.5
            pd /= 1e5 # pa to bar
            pd -= p0 / 1e5
            self.pd2 = -pd
            self.aspen.Tree.FindNode(r"\Data\Blocks\B2\Input\PDROP").value = self.pd2
            print('Pressure Checked!')
            

class SAO():
    def __init__(self):
        self.a = Aspen()
        self.initial_solution = self.a.get_original_params()
        self.current_solution = self.initial_solution
        self.best_solution = self.initial_solution

        self.best_score = 0
        self.temperature = 1000
        self.cooling_rate = 0.7
        self.pause_time = 60 # 1sec
        self.varialbes_num = len(self.current_solution)
        self.press = [77, 90, 104, 120] # 290, 300, 310, 320
        self.lb = [300, 200, 100, 1, 0.5, 1, 0.5]
        self.ub = [400, 2000, 270, 5, 1, 5, 1]
        self.iterations = 10
        self.n = 1
        self.recorder = []
    
    def press_select(self):
        dice = np.random.randint(len(self.press))
        self.current_solution[7] = self.press[dice]

    
    def objective(self):
        """
        Give penalty while:
        The ratio of length to diameter larger than 5
        When the temperature can not be control
        When process temperature is too close to coolant temperature
        """
        self.conv, overshoot, close = self.a.get_result()
        if self.best_score == 0:
            conv_origin = self.a.get_result()[0][2] 
            self.best_conv = max(self.conv[2], conv_origin)
            score = -(self.conv[2] - conv_origin)**2
            if overshoot == 1:
                score += 5e4
            if (self.current_solution[3] / self.current_solution[4]) or (self.current_solution[5] / self.current_solution[6]) > 5:
                score += 5e1
            if close == 1:
                score += 5e2
            self.best_score = score
        else:
            score = -(self.conv[2] - self.best_conv)**2
            if overshoot == 1:
                score += 5e4
            if (self.current_solution[3] / self.current_solution[4]) or (self.current_solution[5] / self.current_solution[6]) > 5:
                score += 5e1     
            if close == 1:
                score += 5e2      
        if self.conv[2] > self.best_conv:
            self.best_conv = self.conv[2]
        return score

    def annealing(self):
        data_label = ['Status', 'block1', 'block2', 'block3', 'reactor1 pressure drop', 'reactor2 pressure drop',
                      'reactor1 conversion', 'reacotr2 conversion', 'overall conversion',
                      'Feed temperature', 'Coolant flowarate', 'Coolant output temperature', 'reactor1 length',
                      'reactor1 diameter', 'reactor2 length', 'reactor2 diameter', 'Coolant saturated pressure']
        block1_type = self.a.aspen.Tree.FindNode(r"\Data\Blocks\B1").AttributeValue(6)
        block2_type = self.a.aspen.Tree.FindNode(r"\Data\Blocks\B2").AttributeValue(6)
        block3_type = self.a.aspen.Tree.FindNode(r"\Data\Blocks\B3").AttributeValue(6)

        score = self.objective()
        xs = [0, 0]
        ys = [self.best_score, self.best_score]
        start = time.time()
        seed = int(start)
        np.random.seed(seed)
        num = 0
        for i in range(9999):
            for j in range(self.iterations):
                num += 1
                for k in range(self.varialbes_num - 1):
                    self.current_solution[k] = np.exp(np.log(self.best_solution[k]) + 0.5*np.random.uniform(-1, 1)) # Steps
                    self.current_solution[k] = max(min(self.current_solution[k], self.ub[k]), self.lb[k])
                self.press_select()

                self.a.set_newparams(self.current_solution)
                status = self.a.check_validation()
                if status == 3:
                    new_score = self.objective()
                    cost = np.abs(score - new_score)
                    if i == 0 and j == 0:
                        overall_cost = cost
                    if new_score > self.best_score:
                        p = np.exp(-cost/(overall_cost*self.temperature))
                        print('passing rate', p)
                        if random.random() < p:
                            accept = True
                        else:
                            accept = False
                    else:
                        accept = True

                    if accept == True:
                        self.best_solution = self.current_solution
                        self.best_score = new_score
                        self.n += 1
                        overall_cost = (overall_cost * (self.n - 1) + cost) / self.n

                        xs[0] = xs[1]
                        ys[0] = ys[1]
                        xs[1] = self.n
                        ys[1] = self.best_score
                        plt.plot(xs, ys, 'b-')
                        plt.xlabel('Number of iterations')
                        plt.ylabel('Process squared error')
                        plt.title('Process improvement', fontsize=10)
                        plt.pause(0.1)
                        plt.savefig('Improvement path')
                        
                        data_value = [status, block1_type, block2_type, block3_type, self.a.pd1, self.a.pd2, self.conv[0],
                                      self.conv[1], self.conv[2]]
                        data_value += self.current_solution
                        self.a.record(data_value, data_label)

                else:
                    data_value = [status, block1_type, block2_type, block3_type, self.a.pd1, self.a.pd2, self.conv[0],
                                  self.conv[1], self.conv[2]]
                    data_value += self.current_solution
                    self.a.record(data_value, data_label)
                    self.a.diversed()
                    break

            print('Accept rate : {:.2f}'.format(self.n/num))
            self.temperature *= self.cooling_rate
            print('Recent Temperature :', self.temperature)
            self.recorder.append(self.best_score)
            if self.temperature < 1:
                break
                
        print('Random seed :', seed)
        end = time.time()
        print('Time consumed : {:.2f} sec'.format(end - start))
        print('Best conversion : {:.2f}'.format(self.best_conv))
        print('Best operation point :', self.best_solution)
        self.a.aspen.Close()



task = SAO()
task.annealing()
var app = angular.module('myApp', ['ngFileUpload']);
app.controller('WordToPdfCtrl', ['$scope', '$document', 'Upload', '$http', function($scope, $document, Upload, $http) {
    function State() {
        var self = this;
        this.preview = false;
        this.actions = false;
        this.loader = false;
        this.files = [];
        this.select = function() {
            var element = $document[0].getElementById('file');
            element.click();
        };
        this.reset = function() {
            $scope.state.preview = false;
            $scope.state.actions = false;
            $scope.state.loader = false;
        }
        this.close = function(element, file) {
            $(element).parent().remove();
            if ($scope.state.files.length === 1) {
                self.reset();
            }
            $scope.state.files.forEach(function(value, index, array) {
                if (file === value) {
                    array.splice(index, 1);
                }
            });
        }
        this.convert = function(files) {
            var data = [];
            files.forEach(function(file, index, array) {
                data.push(file.name);
            });
            $scope.state.loader = true;
            $scope.state.actions = false;
            $http({
                method: 'POST',
                url: '/convert_from_pdf',
                data: data,
                headers: {
                    "Content-Type": "application/json"
                },
                responseType: 'blob',
                dataType: 'json'
            }).then(function(success) {
                $scope.state.download = true;
                $scope.state.loader = false;
                var file = new Blob([success.data], { type: 'application/zip' });
                saveAs(file, 'output.zip');
            }, function(err) {
                $scope.state.loader = false;
                alert('An error occured');
            });
        }
        this.download = function() {
            $scope.state.loader = true;
            $scope.state.actions = false;
            $http({
                method: 'POST',
                url: '/download',
                responseType: 'blob'
            }).then(function(success) {
                $scope.state.loader = false;
                var file = new Blob([success.data], { type: 'application/zip' });
                saveAs(file, 'output.zip');
            }, function(err) {
                $scope.state.loader = false;
                alert('File no longer exist');
            });
        }
        this.uploadFiles = function(files, errFiles) {
            $scope.files = files;
            $scope.errFiles = errFiles;
            angular.forEach(files, function(file) {
                file.upload = Upload.upload({
                    url: $scope.url,
                    data: { file: file }
                });
                file.upload.then(function(success) {
                    //file.result = success.data;
                    if (success.data === "error") {
                        alert("Check your network");
                        $scope.state.loader = false;
                        return;
                    }
                    $scope.state.preview = true;
                    $scope.state.actions = true;
                    $scope.state.loader = false;
                    $scope.state.files.push(file);
                }, function(error) {
                    if (error.status === 404) {
                        alert(error.data);
                        $scope.state.loader = false;
                        return;
                    } else {
                        alert("Check your network");
                        $scope.state.loader = false;
                    }
                }, function(progress) {
                    $scope.state.loader = true;
                });
            });
        };
    }
    $scope.state = new State();
}]);